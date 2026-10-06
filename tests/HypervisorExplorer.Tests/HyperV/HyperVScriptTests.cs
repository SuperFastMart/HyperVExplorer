using System.Text;
using System.Text.RegularExpressions;
using HypervisorExplorer.Collectors.HyperV;

namespace HypervisorExplorer.Tests.HyperV;

public class HyperVScriptTests
{
    public static TheoryData<string, string> Scripts() => new()
    {
        { HyperVScripts.CollectScriptName, HyperVScripts.CollectScript },
        { HyperVScripts.WrapperScriptName, HyperVScripts.WrapperScript },
        { "Bootstrap", HyperVScripts.Bootstrap },
    };

    [Fact]
    public void EmbeddedScripts_Load()
    {
        Assert.Contains("schemaVersion", HyperVScripts.CollectScript);
        Assert.Contains("Hyper-V\\Get-VM", HyperVScripts.CollectScript);
        Assert.Contains("ConvertTo-Json -InputObject $result -Depth 12 -Compress", HyperVScripts.CollectScript);
        Assert.Equal(HyperVScripts.CollectScript, HyperVCollector.GetCollectionScript());
        Assert.Contains("Invoke-Command", HyperVScripts.WrapperScript);
        Assert.Contains("-ThrottleLimit 8", HyperVScripts.WrapperScript);
        Assert.Contains("PROGRESS: ", HyperVScripts.WrapperScript);
    }

    [Fact]
    public void Bootstrap_FitsComfortablyOnTheCommandLine()
    {
        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(HyperVScripts.Bootstrap));
        Assert.True(encoded.Length < 8000, $"Encoded bootstrap is {encoded.Length} chars");
    }

    [Theory]
    [MemberData(nameof(Scripts))]
    public void Scripts_AreWindowsPowerShell51Compatible(string name, string script)
    {
        string[] forbidden = ["??", "?.", "?[", " ? ", "-AsHashtable", "-Parallel", "&&", "||", "$PSStyle", "Get-Error", "-SkipCertificateCheck", "ConvertFrom-Json -Depth"];
        foreach (var token in forbidden)
            Assert.False(script.Contains(token, StringComparison.Ordinal), $"{name} contains PowerShell 7-only token '{token}'");
        Assert.DoesNotMatch(new Regex(@"ConvertFrom-Json[^\r\n]*-Depth"), script);
        Assert.DoesNotMatch(new Regex(@"\bclean\s*\{"), script);
    }

    [Theory]
    [MemberData(nameof(Scripts))]
    public void Scripts_HaveBalancedDelimitersOutsideStringsAndComments(string name, string script)
    {
        var stack = new Stack<(char Ch, int Line)>();
        var line = 1;
        var i = 0;
        while (i < script.Length)
        {
            var c = script[i];
            if (c == '\n') { line++; i++; continue; }
            if (c == '<' && i + 1 < script.Length && script[i + 1] == '#')
            {
                var end = script.IndexOf("#>", i + 2, StringComparison.Ordinal);
                Assert.True(end > 0, $"{name}: unterminated block comment at line {line}");
                line += script[i..end].Count(ch => ch == '\n');
                i = end + 2;
                continue;
            }
            if (c == '#')
            {
                while (i < script.Length && script[i] != '\n') i++;
                continue;
            }
            if (c == '\'')
            {
                var start = line;
                i++;
                while (true)
                {
                    Assert.True(i < script.Length, $"{name}: unterminated single-quoted string from line {start}");
                    if (script[i] == '\n') line++;
                    if (script[i] == '\'')
                    {
                        if (i + 1 < script.Length && script[i + 1] == '\'') { i += 2; continue; }
                        i++;
                        break;
                    }
                    i++;
                }
                continue;
            }
            if (c == '"')
            {
                var start = line;
                i++;
                var depth = 0; // $( ... ) sub-expressions inside the string
                while (true)
                {
                    Assert.True(i < script.Length, $"{name}: unterminated double-quoted string from line {start}");
                    var s = script[i];
                    if (s == '\n') line++;
                    if (s == '`') { i += 2; continue; }
                    if (s == '$' && i + 1 < script.Length && script[i + 1] == '(') { depth++; i += 2; continue; }
                    if (s == ')' && depth > 0) { depth--; i++; continue; }
                    if (s == '"' && depth == 0)
                    {
                        if (i + 1 < script.Length && script[i + 1] == '"') { i += 2; continue; }
                        i++;
                        break;
                    }
                    i++;
                }
                continue;
            }
            if (c == '`') { i += 2; continue; }
            if (c is '(' or '{' or '[') stack.Push((c, line));
            else if (c is ')' or '}' or ']')
            {
                Assert.True(stack.Count > 0, $"{name}: unexpected '{c}' at line {line}");
                var (open, openLine) = stack.Pop();
                var expected = open switch { '(' => ')', '{' => '}', _ => ']' };
                Assert.True(c == expected, $"{name}: '{c}' at line {line} closes '{open}' from line {openLine}");
            }
            i++;
        }
        Assert.True(stack.Count == 0, stack.Count == 0 ? "" : $"{name}: unclosed '{stack.Peek().Ch}' from line {stack.Peek().Line}");
    }

    [Fact]
    public void CollectScript_DoesNotAssignAutomaticVariables()
    {
        foreach (var v in new[] { "$host", "$input", "$args", "$error", "$matches", "$pid", "$profile", "$home" })
        {
            Assert.DoesNotMatch(new Regex(@"(?i)" + Regex.Escape(v) + @"\s*=[^=]"), HyperVScripts.CollectScript);
            Assert.DoesNotMatch(new Regex(@"(?i)" + Regex.Escape(v) + @"\s*=[^=]"), HyperVScripts.WrapperScript);
        }
    }
}
