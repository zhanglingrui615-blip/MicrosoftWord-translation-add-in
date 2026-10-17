using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

namespace WordVocabulary
{
    public static class ExcelExport
    {
        private const string Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static void Part(ZipArchive zip, string name, string xml)
        {
            using (var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false))) writer.Write(xml);
        }
        private static string CellText(string value)
        {
            var text = new StringBuilder();
            value = (value ?? "").Replace("\r\n", "\n").Replace("\r", "\n");
            for (int i = 0; i < value.Length && text.Length < 32767; i++)
            {
                if (XmlConvert.IsXmlChar(value[i])) text.Append(value[i]);
                else if (Char.IsHighSurrogate(value[i]) && i + 1 < value.Length && Char.IsLowSurrogate(value[i + 1]) && text.Length < 32766)
                { text.Append(value[i]); text.Append(value[++i]); }
            }
            return text.ToString();
        }
        private static void Row(XmlWriter writer, int number, string[] values, bool header)
        {
            writer.WriteStartElement("row", Ns); writer.WriteAttributeString("r", number.ToString());
            if (header) { writer.WriteAttributeString("ht", "26"); writer.WriteAttributeString("customHeight", "1"); }
            for (int i = 0; i < values.Length; i++)
            {
                writer.WriteStartElement("c", Ns);
                writer.WriteAttributeString("r", ((char)('A' + i)).ToString() + number);
                writer.WriteAttributeString("t", "inlineStr"); writer.WriteAttributeString("s", header ? "1" : "2");
                writer.WriteStartElement("is", Ns); writer.WriteStartElement("t", Ns);
                writer.WriteAttributeString("xml", "space", null, "preserve");
                // Inline strings preserve formula-like notes as text, never executable formulas.
                writer.WriteString(CellText(values[i]));
                writer.WriteEndElement(); writer.WriteEndElement(); writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }
        public static void Save(string path, IEnumerable<Entry> entries)
        {
            var rows = entries.ToList();
            if (rows.Count > 1048575) throw new InvalidOperationException("生词数量超过 Excel 单张工作表上限。");
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.ReadWrite))
                using (var zip = new ZipArchive(file, ZipArchiveMode.Create))
                {
                    Part(zip, "[Content_Types].xml", "<?xml version='1.0' encoding='utf-8'?><Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/><Override PartName='/xl/styles.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'/><Override PartName='/xl/worksheets/sheet1.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'/></Types>");
                    Part(zip, "_rels/.rels", "<?xml version='1.0' encoding='utf-8'?><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl/workbook.xml'/></Relationships>");
                    Part(zip, "xl/workbook.xml", "<?xml version='1.0' encoding='utf-8'?><workbook xmlns='" + Ns + "' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets><sheet name='生词本' sheetId='1' r:id='rId1'/></sheets></workbook>");
                    Part(zip, "xl/_rels/workbook.xml.rels", "<?xml version='1.0' encoding='utf-8'?><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/sheet1.xml'/><Relationship Id='rId2' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/></Relationships>");
                    Part(zip, "xl/styles.xml", "<?xml version='1.0' encoding='utf-8'?><styleSheet xmlns='" + Ns + "'><fonts count='2'><font><sz val='11'/><name val='Calibri'/></font><font><b/><sz val='11'/><color rgb='FFFFFFFF'/><name val='Calibri'/></font></fonts><fills count='3'><fill><patternFill patternType='none'/></fill><fill><patternFill patternType='gray125'/></fill><fill><patternFill patternType='solid'><fgColor rgb='FF2E75B6'/><bgColor indexed='64'/></patternFill></fill></fills><borders count='1'><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count='1'><xf numFmtId='0' fontId='0' fillId='0' borderId='0'/></cellStyleXfs><cellXfs count='3'><xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/><xf numFmtId='0' fontId='1' fillId='2' borderId='0' xfId='0' applyAlignment='1'><alignment horizontal='center' vertical='center' wrapText='1'/></xf><xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0' applyAlignment='1'><alignment vertical='top' wrapText='1'/></xf></cellXfs><cellStyles count='1'><cellStyle name='Normal' xfId='0' builtinId='0'/></cellStyles></styleSheet>");
                    using (var stream = zip.CreateEntry("xl/worksheets/sheet1.xml").Open())
                    using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), NewLineHandling = NewLineHandling.None }))
                    {
                        writer.WriteStartDocument(); writer.WriteStartElement("worksheet", Ns);
                        writer.WriteStartElement("dimension", Ns); writer.WriteAttributeString("ref", "A1:D" + (rows.Count + 1)); writer.WriteEndElement();
                        writer.WriteStartElement("sheetViews", Ns); writer.WriteStartElement("sheetView", Ns); writer.WriteAttributeString("workbookViewId", "0");
                        writer.WriteStartElement("pane", Ns); writer.WriteAttributeString("ySplit", "1"); writer.WriteAttributeString("topLeftCell", "A2"); writer.WriteAttributeString("activePane", "bottomLeft"); writer.WriteAttributeString("state", "frozen"); writer.WriteEndElement();
                        writer.WriteEndElement(); writer.WriteEndElement();
                        writer.WriteStartElement("cols", Ns);
                        int[] widths = { 25, 70, 40, 30 };
                        for (int i = 0; i < widths.Length; i++) { writer.WriteStartElement("col", Ns); writer.WriteAttributeString("min", (i + 1).ToString()); writer.WriteAttributeString("max", (i + 1).ToString()); writer.WriteAttributeString("width", widths[i].ToString()); writer.WriteAttributeString("customWidth", "1"); writer.WriteEndElement(); }
                        writer.WriteEndElement(); writer.WriteStartElement("sheetData", Ns);
                        Row(writer, 1, new[] { "单词", "原句", "个人笔记", "收藏时间" }, true);
                        for (int i = 0; i < rows.Count; i++) Row(writer, i + 2, new[] { rows[i].Word, rows[i].Sentence, rows[i].Note, rows[i].Created }, false);
                        writer.WriteEndElement();
                        writer.WriteStartElement("autoFilter", Ns); writer.WriteAttributeString("ref", "A1:D" + (rows.Count + 1)); writer.WriteEndElement();
                        writer.WriteEndElement(); writer.WriteEndDocument();
                    }
                }
                if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
