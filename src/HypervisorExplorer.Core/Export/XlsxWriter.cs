using System.Globalization;
using System.IO.Compression;
using System.Security;
using System.Text;

namespace HypervisorExplorer.Core.Export;

/// <summary>A worksheet to write: header row plus typed cell values.</summary>
public sealed class XlsxSheet
{
    public required string Name { get; init; }
    public required IReadOnlyList<string> Headers { get; init; }
    public required IReadOnlyList<object?[]> Rows { get; init; }
    public int FrozenColumns { get; init; } = 1;
    public bool AutoFilter { get; init; } = true;
}

/// <summary>
/// Minimal, dependency-free SpreadsheetML (.xlsx) writer. Supports shared strings, numbers, dates,
/// bold header, frozen panes, autofilter and content-based column widths -- enough for RVTools-style workbooks.
/// </summary>
public static class XlsxWriter
{
    private const int MaxCellText = 32767;
    private const int StyleHeader = 1;
    private const int StyleDate = 2;
    private const int StyleDecimal = 3;

    public static void Write(string path, IReadOnlyList<XlsxSheet> sheets)
    {
        var dir = Path.GetDirectoryName(Path.GetFullPath(path));
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
        var temp = path + ".tmp";
        using (var fs = new FileStream(temp, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            Write(fs, sheets);
        }
        File.Move(temp, path, overwrite: true);
    }

    public static void Write(Stream output, IReadOnlyList<XlsxSheet> sheets)
    {
        if (sheets.Count == 0) throw new ArgumentException("At least one sheet is required.", nameof(sheets));

        var names = UniqueSheetNames(sheets.Select(s => s.Name));
        var strings = new SharedStrings();

        using var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);

        for (var i = 0; i < sheets.Count; i++)
        {
            WriteEntry(zip, $"xl/worksheets/sheet{i + 1}.xml", w => WriteSheet(w, sheets[i], strings));
        }

        WriteEntry(zip, "[Content_Types].xml", w => w.Write(ContentTypes(sheets.Count)));
        WriteEntry(zip, "_rels/.rels", w => w.Write(RootRels));
        WriteEntry(zip, "docProps/core.xml", w => w.Write(CoreProps()));
        WriteEntry(zip, "docProps/app.xml", w => w.Write(AppProps));
        WriteEntry(zip, "xl/workbook.xml", w => w.Write(Workbook(sheets, names)));
        WriteEntry(zip, "xl/_rels/workbook.xml.rels", w => w.Write(WorkbookRels(sheets.Count)));
        WriteEntry(zip, "xl/styles.xml", w => w.Write(Styles));
        WriteEntry(zip, "xl/sharedStrings.xml", strings.WriteTo);
    }

    private static void WriteEntry(ZipArchive zip, string name, Action<TextWriter> write)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        write(writer);
    }

    private static void WriteSheet(TextWriter w, XlsxSheet sheet, SharedStrings strings)
    {
        var colCount = sheet.Headers.Count;
        var widths = sheet.Headers.Select(h => (double)Math.Min(h.Length + 4, 60)).ToArray();
        foreach (var row in sheet.Rows.Take(500))
        {
            for (var c = 0; c < colCount && c < row.Length; c++)
            {
                var len = row[c] switch
                {
                    null => 0,
                    DateTime => 18,
                    var v => HypervisorExplorer.Core.Tables.TableColumn.Format(v).Length,
                };
                widths[c] = Math.Min(Math.Max(widths[c], len + 2), 60);
            }
        }

        w.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        w.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        var lastRef = colCount > 0 ? $"{ColumnName(colCount - 1)}{sheet.Rows.Count + 1}" : "A1";
        w.Write($"<dimension ref=\"A1:{lastRef}\"/>");

        w.Write("<sheetViews><sheetView workbookViewId=\"0\">");
        if (sheet.FrozenColumns > 0 || colCount > 0)
        {
            var x = Math.Min(sheet.FrozenColumns, Math.Max(colCount - 1, 0));
            var topLeft = $"{ColumnName(x)}2";
            w.Write(x > 0
                ? $"<pane xSplit=\"{x}\" ySplit=\"1\" topLeftCell=\"{topLeft}\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\" activeCell=\"{topLeft}\" sqref=\"{topLeft}\"/>"
                : "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\" activeCell=\"A2\" sqref=\"A2\"/>");
        }
        w.Write("</sheetView></sheetViews>");
        w.Write("<sheetFormatPr defaultRowHeight=\"15\"/>");

        if (colCount > 0)
        {
            w.Write("<cols>");
            for (var c = 0; c < colCount; c++)
                w.Write($"<col min=\"{c + 1}\" max=\"{c + 1}\" width=\"{widths[c].ToString("0.#", CultureInfo.InvariantCulture)}\" customWidth=\"1\"/>");
            w.Write("</cols>");
        }

        w.Write("<sheetData>");
        w.Write("<row r=\"1\">");
        for (var c = 0; c < colCount; c++)
            w.Write($"<c r=\"{ColumnName(c)}1\" t=\"s\" s=\"{StyleHeader}\"><v>{strings.Index(sheet.Headers[c])}</v></c>");
        w.Write("</row>");

        var r = 2;
        foreach (var row in sheet.Rows)
        {
            w.Write($"<row r=\"{r}\">");
            for (var c = 0; c < colCount && c < row.Length; c++)
                WriteCell(w, $"{ColumnName(c)}{r}", row[c], strings);
            w.Write("</row>");
            r++;
        }
        w.Write("</sheetData>");

        if (sheet.AutoFilter && colCount > 0)
            w.Write($"<autoFilter ref=\"A1:{lastRef}\"/>");

        w.Write("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>");
        w.Write("</worksheet>");
    }

    private static void WriteCell(TextWriter w, string cellRef, object? value, SharedStrings strings)
    {
        switch (value)
        {
            case null:
                return;
            case DateTime dt:
                w.Write($"<c r=\"{cellRef}\" s=\"{StyleDate}\"><v>{dt.ToOADate().ToString("R", CultureInfo.InvariantCulture)}</v></c>");
                return;
            case DateTimeOffset dto:
                WriteCell(w, cellRef, dto.LocalDateTime, strings);
                return;
            case byte or sbyte or short or ushort or int or uint or long or ulong:
                w.Write($"<c r=\"{cellRef}\"><v>{Convert.ToString(value, CultureInfo.InvariantCulture)}</v></c>");
                return;
            case double d when double.IsFinite(d):
                w.Write($"<c r=\"{cellRef}\"{(d % 1 != 0 ? $" s=\"{StyleDecimal}\"" : "")}><v>{d.ToString("R", CultureInfo.InvariantCulture)}</v></c>");
                return;
            case float f when float.IsFinite(f):
                WriteCell(w, cellRef, (double)f, strings);
                return;
            case decimal m:
                WriteCell(w, cellRef, (double)m, strings);
                return;
            case bool b:
                WriteCell(w, cellRef, b ? "True" : "False", strings);
                return;
            default:
                var text = HypervisorExplorer.Core.Tables.TableColumn.Format(value);
                if (text.Length == 0) return;
                if (text.Length > MaxCellText) text = text[..MaxCellText];
                w.Write($"<c r=\"{cellRef}\" t=\"s\"><v>{strings.Index(text)}</v></c>");
                return;
        }
    }

    public static string ColumnName(int index)
    {
        var sb = new StringBuilder();
        index++;
        while (index > 0)
        {
            var rem = (index - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            index = (index - 1) / 26;
        }
        return sb.ToString();
    }

    public static List<string> UniqueSheetNames(IEnumerable<string> names)
    {
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<string>();
        foreach (var raw in names)
        {
            var clean = new string(raw.Select(ch => "[]:*?/\\".Contains(ch) ? '_' : ch).ToArray()).Trim('\'');
            if (clean.Length == 0) clean = "Sheet";
            if (clean.Length > 31) clean = clean[..31];
            var candidate = clean;
            for (var n = 2; !used.Add(candidate); n++)
            {
                var suffix = $" ({n})";
                candidate = (clean.Length + suffix.Length > 31 ? clean[..(31 - suffix.Length)] : clean) + suffix;
            }
            result.Add(candidate);
        }
        return result;
    }

    internal static string Escape(string s)
    {
        var sb = new StringBuilder(s.Length);
        foreach (var ch in s)
        {
            // Drop characters that are illegal in XML 1.0 (control chars other than tab/CR/LF).
            if (ch < 0x20 && ch != '\t' && ch != '\n' && ch != '\r') continue;
            if (ch is '￾' or '￿') continue;
            sb.Append(ch);
        }
        return SecurityElement.Escape(sb.ToString()) ?? "";
    }

    private sealed class SharedStrings
    {
        private readonly Dictionary<string, int> _index = new(StringComparer.Ordinal);
        private readonly List<string> _values = [];
        private int _count;

        public int Index(string s)
        {
            _count++;
            if (_index.TryGetValue(s, out var i)) return i;
            i = _values.Count;
            _values.Add(s);
            _index[s] = i;
            return i;
        }

        public void WriteTo(TextWriter w)
        {
            w.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            w.Write($"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{_count}\" uniqueCount=\"{_values.Count}\">");
            foreach (var v in _values)
            {
                var escaped = Escape(v);
                var preserve = v.Length > 0 && (char.IsWhiteSpace(v[0]) || char.IsWhiteSpace(v[^1])) ? " xml:space=\"preserve\"" : "";
                w.Write($"<si><t{preserve}>{escaped}</t></si>");
            }
            w.Write("</sst>");
        }
    }

    private static string ContentTypes(int sheetCount)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
        for (var i = 1; i <= sheetCount; i++)
            sb.Append($"<Override PartName=\"/xl/worksheets/sheet{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
        sb.Append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
        sb.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
        sb.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
        sb.Append("</Types>");
        return sb.ToString();
    }

    private const string RootRels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" +
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
        "</Relationships>";

    private const string AppProps =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>Hypervisor Explorer</Application></Properties>";

    private static string CoreProps()
    {
        var now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
               "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
               "<dc:creator>Hypervisor Explorer</dc:creator>" +
               $"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:created>" +
               $"<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:modified>" +
               "</cp:coreProperties>";
    }

    private static string Workbook(IReadOnlyList<XlsxSheet> sheets, IReadOnlyList<string> names)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        sb.Append("<bookViews><workbookView/></bookViews><sheets>");
        for (var i = 0; i < names.Count; i++)
            sb.Append($"<sheet name=\"{Escape(names[i])}\" sheetId=\"{i + 1}\" r:id=\"rId{i + 1}\"/>");
        sb.Append("</sheets>");

        var filters = sheets.Select((s, i) => (s, i)).Where(t => t.s.AutoFilter && t.s.Headers.Count > 0).ToList();
        if (filters.Count > 0)
        {
            sb.Append("<definedNames>");
            foreach (var (s, i) in filters)
            {
                var sheetRef = "'" + names[i].Replace("'", "''") + "'";
                sb.Append($"<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"{i}\" hidden=\"1\">{Escape(sheetRef)}!$A$1:${ColumnName(s.Headers.Count - 1)}${s.Rows.Count + 1}</definedName>");
            }
            sb.Append("</definedNames>");
        }
        sb.Append("</workbook>");
        return sb.ToString();
    }

    private static string WorkbookRels(int sheetCount)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        for (var i = 1; i <= sheetCount; i++)
            sb.Append($"<Relationship Id=\"rId{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i}.xml\"/>");
        sb.Append($"<Relationship Id=\"rId{sheetCount + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
        sb.Append($"<Relationship Id=\"rId{sheetCount + 2}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
        sb.Append("</Relationships>");
        return sb.ToString();
    }

    // Styles: 0 default, 1 bold header with fill + bottom border, 2 date/time, 3 two-decimal number.
    private const string Styles =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
        "<numFmts count=\"2\"><numFmt numFmtId=\"164\" formatCode=\"yyyy\\-mm\\-dd\\ hh:mm:ss\"/><numFmt numFmtId=\"165\" formatCode=\"0.00\"/></numFmts>" +
        "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font><font><b/><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>" +
        "<fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>" +
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFDCE6F1\"/><bgColor indexed=\"64\"/></patternFill></fill></fills>" +
        "<borders count=\"2\"><border><left/><right/><top/><bottom/><diagonal/></border>" +
        "<border><left/><right/><top/><bottom style=\"thin\"><color rgb=\"FF95B3D7\"/></bottom><diagonal/></border></borders>" +
        "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
        "<cellXfs count=\"4\">" +
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>" +
        "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
        "<xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
        "</cellXfs>" +
        "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
        "</styleSheet>";
}
