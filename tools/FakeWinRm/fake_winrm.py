#!/usr/bin/env python3
"""
Fake WinRM endpoint for testing Hypervisor Explorer's built-in WinRM client without a Windows host.

It authenticates with NTLM using pyspnego (the library behind pywinrm/Ansible's Windows support), seals and
unseals bodies with WinRM's HTTP-SPNEGO-session-encrypted framing, and implements the Windows Remote Shell
operations the collector uses. Instead of running PowerShell, it answers the collector's bootstrap requests
with canned Collect-HyperV.ps1 output, so the whole Mac-side pipeline is exercised end to end.

    pip install pyspnego
    python3 fake_winrm.py --port 15985 --node1 hv-node1.json --node2 hv-node2.json

Accounts (NTLM): CONTOSO\\admin / Passw0rd!  (full access), CONTOSO\\nonadmin / Passw0rd! (shell access denied).
Connect as "localhost" for node HV01 (the primary) and "127.0.0.1" for node HV02.
"""
import argparse
import base64
import gzip
import json
import os
import re
import struct
import sys
import tempfile
import threading
import uuid
import xml.etree.ElementTree as ET
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import spnego

BOUNDARY = b"--Encrypted Boundary"
PROTOCOL = b"application/HTTP-SPNEGO-session-encrypted"
CT_ENCRYPTED = 'multipart/encrypted;protocol="application/HTTP-SPNEGO-session-encrypted";boundary="Encrypted Boundary"'
NS = {
    "s": "http://www.w3.org/2003/05/soap-envelope",
    "a": "http://schemas.xmlsoap.org/ws/2004/08/addressing",
    "w": "http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd",
    "rsp": "http://schemas.microsoft.com/wbem/wsman/1/windows/shell",
}
ENV_OPEN = ('<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" '
            'xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing" '
            'xmlns:w="http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd" '
            'xmlns:rsp="http://schemas.microsoft.com/wbem/wsman/1/windows/shell"><s:Header/><s:Body>')
ENV_CLOSE = "</s:Body></s:Envelope>"

STATE = {"shells": {}, "errors": [], "requests": 0}
LOCK = threading.Lock()
NODES = {}


def fail(msg):
    with LOCK:
        STATE["errors"].append(msg)
    print("CHECK FAILED: " + msg, file=sys.stderr, flush=True)


def fault(code, subcode, text):
    return (ENV_OPEN + '<s:Fault><s:Code><s:Value>s:Receiver</s:Value><s:Subcode><s:Value>' + subcode +
            '</s:Value></s:Subcode></s:Code><s:Reason><s:Text xml:lang="en-US">' + text +
            '</s:Text></s:Reason><s:Detail><f:WSManFault xmlns:f="http://schemas.microsoft.com/wbem/wsman/1/wsmanfault" Code="' +
            code + '" Machine="fake"><f:Message>' + text + '</f:Message></f:WSManFault></s:Detail></s:Fault>' + ENV_CLOSE)


def gz_b64(text):
    return base64.b64encode(gzip.compress(text.encode("utf-8"))).decode()


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    ctx = None
    user = None

    def log_message(self, fmt, *args):
        pass

    def reply(self, status, body=b"", headers=None):
        self.send_response(status)
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    # ------------------------------------------------------------ framing (mirrors pywinrm's encryption.py)
    def unseal(self, body):
        parts = [p for p in re.split(re.escape(BOUNDARY) + rb"\r\n", body) if p]
        header, payload = parts[0], parts[1]
        expected = int(header.split(b"Length=")[1].split(b"\r\n")[0])
        payload = payload.replace(b"\tContent-Type: application/octet-stream\r\n", b"", 1)
        if payload.endswith(BOUNDARY + b"--\r\n"):
            payload = payload[: -len(BOUNDARY + b"--\r\n")]
        sig_len = struct.unpack("<i", payload[:4])[0]
        sig, data = payload[4:4 + sig_len], payload[4 + sig_len:]
        plain = self.ctx.unwrap_winrm(sig, data)
        if len(plain) != expected:
            fail(f"OriginalContent Length={expected} but decrypted {len(plain)} bytes")
        return plain

    def seal(self, plain):
        res = self.ctx.wrap_winrm(plain)
        msg = BOUNDARY + b"\r\n\tContent-Type: " + PROTOCOL + b"\r\n"
        msg += b"\tOriginalContent: type=application/soap+xml;charset=UTF-8;Length=" + str(len(plain)).encode() + b"\r\n"
        msg += BOUNDARY + b"\r\n\tContent-Type: application/octet-stream\r\n"
        msg += struct.pack("<i", len(res.header)) + res.header + res.data
        msg += BOUNDARY + b"--\r\n"
        return msg

    # ------------------------------------------------------------ HTTP
    def do_POST(self):
        with LOCK:
            STATE["requests"] += 1
        body = self.rfile.read(int(self.headers.get("Content-Length", 0)))
        auth = self.headers.get("Authorization")
        if auth:
            scheme, token = auth.split(" ", 1)
            if scheme not in ("Negotiate", "NTLM"):
                return self.reply(401, headers={"WWW-Authenticate": "Negotiate"})
            if self.ctx is None or self.ctx.complete:
                self.ctx = spnego.server(protocol="ntlm")
            try:
                out = self.ctx.step(base64.b64decode(token))
            except Exception as e:  # bad password etc.
                print(f"auth rejected: {e}", file=sys.stderr, flush=True)
                self.ctx = None
                return self.reply(401, headers={"WWW-Authenticate": "Negotiate"})
            if not self.ctx.complete:
                return self.reply(401, headers={"WWW-Authenticate": "Negotiate " + base64.b64encode(out).decode()})
            self.user = self.ctx.client_principal
            print(f"authenticated {self.user} via {self.headers.get('Host')}", file=sys.stderr, flush=True)
            if not body:
                return self.reply(200)

        if self.ctx is None or not self.ctx.complete:
            return self.reply(401, headers={"WWW-Authenticate": "Negotiate"})
        if "multipart/encrypted" not in self.headers.get("Content-Type", ""):
            fail("received an unencrypted message on HTTP")
            return self.reply(400)

        soap = self.unseal(body).decode("utf-8")
        status, resp = self.handle_soap(soap)
        self.reply(status, self.seal(resp.encode("utf-8")), {"Content-Type": CT_ENCRYPTED})

    # ------------------------------------------------------------ WS-Man / WinRS
    def handle_soap(self, soap):
        doc = ET.fromstring(soap)
        action = doc.find(".//a:Action", NS).text
        if doc.find(".//w:ResourceURI", NS).text != "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd":
            fail("unexpected ResourceURI")
        sel = doc.find(".//w:Selector[@Name='ShellId']", NS)
        host = self.headers.get("Host", "").split(":")[0]
        op = action.rsplit("/", 1)[-1]

        if op == "Create":
            if self.user and "nonadmin" in self.user.lower():
                return 500, fault("5", "w:AccessDenied", "Access is denied.")
            sid = str(uuid.uuid4()).upper()
            with LOCK:
                STATE["shells"][sid] = {"stdin": b"", "receives": 0, "host": host, "ended": False}
            return 200, ENV_OPEN + f"<rsp:Shell><rsp:ShellId>{sid}</rsp:ShellId></rsp:Shell>" + ENV_CLOSE

        shell = STATE["shells"].get(sel.text if sel is not None else "")
        if shell is None and op != "Delete":
            fail(f"{op} for unknown shell")
            return 500, fault("2150858843", "w:InvalidSelectors", "The shell was not found.")

        if op == "Command":
            cmd = doc.find(".//rsp:Command", NS).text
            args = doc.find(".//rsp:Arguments", NS).text
            if cmd != "powershell.exe" or "-EncodedCommand" not in args:
                fail(f"unexpected command line: {cmd} {args[:80]}")
            boot = base64.b64decode(args.split("-EncodedCommand ")[1]).decode("utf-16-le")
            if "<<<HVE:RESULT>>>" not in boot:
                fail("bootstrap does not look like the collector bootstrap")
            shell["cid"] = str(uuid.uuid4()).upper()
            return 200, ENV_OPEN + f"<rsp:CommandResponse><rsp:CommandId>{shell['cid']}</rsp:CommandId></rsp:CommandResponse>" + ENV_CLOSE

        if op == "Send":
            st = doc.find(".//rsp:Stream", NS)
            if st.get("CommandId") != shell.get("cid"):
                fail("Send with wrong CommandId")
            shell["stdin"] += base64.b64decode(st.text or "")
            if st.get("End") == "true":
                shell["ended"] = True
            return 200, ENV_OPEN + "<rsp:SendResponse/>" + ENV_CLOSE

        if op == "Receive":
            if not shell["ended"]:
                fail("Receive before stdin was closed")
            shell["receives"] += 1
            n = shell["receives"]
            cid = shell["cid"]
            req = json.loads(base64.b64decode(re.sub(rb"[^A-Za-z0-9+/=]", b"", shell["stdin"])))
            node = NODES.get(host)
            if node is None:
                fail(f"no node data for host {host}")
                node = NODES["localhost"]
            if n == 1:
                prog = f"PROGRESS: {node['name']}: Collecting virtual machines\r\nPROGRESS: {node['name']}: Collecting"
                return 200, ENV_OPEN + f'<rsp:ReceiveResponse><rsp:Stream Name="stderr" CommandId="{cid}">' + \
                    base64.b64encode(prog.encode()).decode() + "</rsp:Stream></rsp:ReceiveResponse>" + ENV_CLOSE
            if n == 2:
                return 500, fault("2150858793", "w:TimedOut",
                                  "The WS-Management service cannot complete the operation within the time specified in OperationTimeout.")
            if req.get("mode") == "discover":
                result = ("PRIMARY|HV01|hv01.contoso.invalid\nNODE|HV01|Up|hv01.contoso.invalid|127.0.0.1\n"
                          "NODE|HV02|Up|hv02.contoso.invalid|127.0.0.1\nNODE|HV03|Down|hv03.contoso.invalid|")
            else:
                if req.get("clusterPrimary") != "HV01":
                    fail(f"clusterPrimary was {req.get('clusterPrimary')!r}")
                if "Collect-HyperV" not in (req.get("script") or ""):
                    fail("collection script missing from request")
                result = node["json"]
            out = " rest of\r\n" + "<<<HVE:RESULT>>>" + gz_b64(result) + "\r\n"
            return 200, ENV_OPEN + f'<rsp:ReceiveResponse><rsp:Stream Name="stderr" CommandId="{cid}">' + \
                base64.b64encode(b" step\r\n").decode() + f'</rsp:Stream><rsp:Stream Name="stdout" CommandId="{cid}">' + \
                base64.b64encode(out.encode()).decode() + f'</rsp:Stream><rsp:Stream Name="stdout" CommandId="{cid}" End="true"></rsp:Stream>' + \
                f'<rsp:CommandState CommandId="{cid}" State="http://schemas.microsoft.com/wbem/wsman/1/windows/shell/CommandState/Done">' + \
                "<rsp:ExitCode>0</rsp:ExitCode></rsp:CommandState></rsp:ReceiveResponse>" + ENV_CLOSE

        if op == "Signal":
            return 200, ENV_OPEN + "<rsp:SignalResponse/>" + ENV_CLOSE

        if op == "Delete":
            with LOCK:
                STATE["shells"].pop(sel.text if sel is not None else "", None)
            return 200, ENV_OPEN + ENV_CLOSE

        fail(f"unsupported action {action}")
        return 500, fault("2150858817", "w:ActionNotSupported", "Unsupported action.")

    def do_GET(self):
        # /status for the test harness
        body = json.dumps({"errors": STATE["errors"], "openShells": len(STATE["shells"]),
                           "requests": STATE["requests"]}).encode()
        self.reply(200, body, {"Content-Type": "application/json"})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=15985)
    ap.add_argument("--node1", required=True)
    ap.add_argument("--node2", required=True)
    a = ap.parse_args()

    users = tempfile.NamedTemporaryFile("w", delete=False, suffix=".txt")
    users.write("CONTOSO:admin:Passw0rd!\nCONTOSO:nonadmin:Passw0rd!\n")
    users.close()
    os.environ["NTLM_USER_FILE"] = users.name

    for host, path, name in (("localhost", a.node1, "HV01"), ("127.0.0.1", a.node2, "HV02")):
        with open(path) as f:
            NODES[host] = {"json": f.read(), "name": name}

    srv = ThreadingHTTPServer(("127.0.0.1", a.port), Handler)
    print(f"fake WinRM listening on 127.0.0.1:{a.port}", flush=True)
    srv.serve_forever()


if __name__ == "__main__":
    main()
