using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using WordVocabulary;
class BookAndExportTests
{
    static int count;
    static void Check(bool ok, string message) { if (!ok) throw new Exception(message); count++; Console.WriteLine("PASS " + message); }
    [STAThread] static int Main(string[] args)
    {
        try { Run(args); return 0; } catch (Exception ex) { Console.WriteLine(ex.GetType().Name + ": " + ex.Message); return 1; }
    }
    static void Run(string[] args)
    {
        string folder = args[0];
        var saved = new List<Entry>();
        int saves = 0;
        int removals = 0;
        Func<Entry> lookup = delegate { return Storage.FindEntry(saved, "bank"); };
        Action remove = delegate { removals++; saved.Clear(); };
        using (var card = new WordCard("bank", "", delegate(string note) { saves++; saved.Add(new Entry { Word = "bank", Note = note }); }, false, lookup, remove))
        {
            card.Location = new Point(-3000, -3000); card.Show(); Application.DoEvents();
            var add = (Button)card.Controls["collect"];
            ((TextBox)card.Controls["note"]).Text = "我的银行笔记";
            add.PerformClick();
            Check(saves == 1 && add.Enabled && add.Text == "已加入生词本", "joined word remains clickable for removal");
            add.PerformClick();
            Check(removals == 1 && saved.Count == 0 && add.Text == "加入生词本" && !((TextBox)card.Controls["note"]).ReadOnly, "second click silently removes and returns to join state");
            add.PerformClick();
            Check(saves == 2 && saved.Count == 1 && add.Text == "已加入生词本", "third click can join again");
            card.Close();
        }
        using (var reopened = new WordCard("BANK", "", delegate(string note) { saves++; }, false, lookup, remove))
        {
            Check(reopened.Controls["collect"].Enabled && reopened.Controls["collect"].Text == "已加入生词本", "reopening existing word shows actionable saved state");
            Check(reopened.Controls["note"].Text == "我的银行笔记", "existing personal note retained");
            reopened.Location = new Point(-3000, -3000); reopened.Show(); Application.DoEvents();
            ((Button)reopened.Controls["collect"]).PerformClick();
            Check(removals == 2 && saved.Count == 0, "reopened saved word can be removed immediately");
            reopened.Close();
        }
        Check(Storage.FindEntry(new[] { new Entry { Word = "campus-only" } }, " CAMPUS\u2011ONLY ") != null, "saved-word matching normalizes case whitespace and hyphens");
        var rows = new List<Entry> { new Entry { Word = "bank", Sentence = "An English sentence.\n第二行：<中文> & quotes \"test\"", Note = "=1+1", Created = "2026-10-17T23:00:00+08:00" }, new Entry { Word = "campus-only", Sentence = "Campus-only access.", Note = "仅限校内访问", Created = "2026-10-17T23:01:00+08:00" } };
        using (var book = new BookForm(delegate { return rows; }))
        {
            book.Location = new Point(-3000, -3000); book.Show(); Application.DoEvents();
            var grid = book.Controls[0].Controls.OfType<DataGridView>().Single();
            int textHeight = TextRenderer.MeasureText("收藏时间", grid.Font).Height;
            Check(book.Text == "生词本" && grid.ColumnHeadersHeight >= textHeight + grid.ColumnHeadersDefaultCellStyle.Padding.Vertical, "book title and padded column headers have enough height");
            using (var image = new Bitmap(book.Width, book.Height)) { book.DrawToBitmap(image, new Rectangle(Point.Empty, book.Size)); image.Save(Path.Combine(folder, "book-preview.png")); }
            book.Close();
        }
        string path = Path.Combine(folder, "export-test.xlsx");
        ExcelExport.Save(path, rows);
        using (var file = File.OpenRead(path))
        using (var zip = new ZipArchive(file, ZipArchiveMode.Read))
        {
            var sheet = new XmlDocument(); using (var stream = zip.GetEntry("xl/worksheets/sheet1.xml").Open()) sheet.Load(stream);
            var ns = new XmlNamespaceManager(sheet.NameTable); ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            Check(sheet.SelectSingleNode("//s:c[@r='B2']/s:is/s:t", ns).InnerText == rows[0].Sentence, "xlsx preserves Chinese newlines quotes and XML characters");
            Check(sheet.SelectSingleNode("//s:c[@r='C2']", ns).Attributes["t"].Value == "inlineStr" && sheet.SelectNodes("//s:f", ns).Count == 0, "formula-like notes exported as text");
            Check(sheet.SelectSingleNode("//s:pane", ns).Attributes["state"].Value == "frozen" && sheet.SelectSingleNode("//s:autoFilter", ns) != null, "xlsx freezes header and enables filters");
        }
        string empty = Path.Combine(folder, "empty-export-test.xlsx"); ExcelExport.Save(empty, new Entry[0]);
        Check(File.Exists(empty), "empty book exports a valid header-only workbook");
        if (args.Length > 1 && args[1] == "--excel") VerifyExcel(path);
        Console.WriteLine(count + " book/export checks passed.");
    }
    static void VerifyExcel(string path)
    {
        dynamic app = null, workbook = null;
        bool owned = false;
        try
        {
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            owned = (int)app.Workbooks.Count == 0;
            app.DisplayAlerts = false;
            workbook = app.Workbooks.Open(Path.GetFullPath(path), 0, true);
            dynamic sheet = workbook.Worksheets[1];
            Check((string)sheet.Cells[1, 1].Value2 == "单词" && (string)sheet.Cells[3, 1].Value2 == "campus-only", "real Excel opens exported workbook with headers and data");
            Check(!(bool)sheet.Cells[2, 3].HasFormula && (string)sheet.Cells[2, 3].Value2 == "=1+1", "real Excel treats note as literal text");
        }
        finally
        {
            try { if (workbook != null) workbook.Close(false); } catch { }
            try { if (app != null && owned && (int)app.Workbooks.Count == 0) app.Quit(); } catch { }
        }
    }
}
