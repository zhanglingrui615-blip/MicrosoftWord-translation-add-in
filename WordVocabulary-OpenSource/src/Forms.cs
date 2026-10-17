using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WordVocabulary
{
    public class WordCard : Form
    {
        private TextBox meaning;
        private Label phonetic;
        private readonly string queriedText;
        private dynamic speaker;
        private readonly bool autoRead;
        private bool autoReadStarted;
        private Button collect;
        private TextBox note;
        private readonly Func<Entry> findSaved;
        protected override bool ShowWithoutActivation { get { return true; } }
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            RefreshSavedState();
            if (autoRead && !autoReadStarted && ReadingInteraction.IsWord(queriedText))
            {
                autoReadStarted = true;
                BeginInvoke(new Action(delegate { if (!IsDisposed) Speak(false); }));
            }
        }
        public WordCard(string word, string sentence, Action<string> save, bool autoRead = true, Func<Entry> findSaved = null, Action remove = null)
        {
            this.findSaved = findSaved ?? delegate { return Storage.FindEntry(Storage.Load(), word); };
            this.autoRead = autoRead;
            queriedText = word;
            Text = word.Length > 45 ? word.Substring(0, 45).Replace("\r", " ").Replace("\n", " ") + "…" : word;
            ClientSize = new Size(420, 170);
            Font = new Font("Microsoft YaHei UI", 10); BackColor = Color.White;
            FormBorderStyle = FormBorderStyle.FixedToolWindow; StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            phonetic = new Label { Text = ReadingInteraction.IsWord(word) ? word : "选中文本翻译", Location = new Point(10, 4), Size = new Size(400, 40), AutoEllipsis = true, ForeColor = Color.FromArgb(35, 45, 55) }; Controls.Add(phonetic);
            meaning = new TextBox { Text = "正在翻译…", Location = new Point(10, 47), Size = new Size(400, 73), ReadOnly = true, Multiline = true, ScrollBars = ScrollBars.Vertical, BackColor = Color.White, BorderStyle = BorderStyle.None }; Controls.Add(meaning);
            Controls.Add(new Label { Text = "笔记", Location = new Point(10, 137), Size = new Size(38, 22), ForeColor = Color.DimGray });
            note = new TextBox { Name = "note", Location = new Point(48, 133), Size = new Size(148, 26) }; Controls.Add(note);
            collect = new Button { Name = "collect", Text = "加入生词本", Location = new Point(204, 131), Size = new Size(110, 29) }; Controls.Add(collect);
            collect.Click += delegate
            {
                try
                {
                    if (this.findSaved() != null) { (remove ?? delegate { Storage.Remove(word); })(); SetSavedState(false); }
                    else { save(note.Text.Trim()); SetSavedState(true); }
                    RefreshSavedState();
                }
                catch (Exception ex) { SetMessage("操作未完成：" + ex.Message); }
            };
            var speak = new Button { Text = "朗读", Location = new Point(322, 131), Size = new Size(86, 29) }; Controls.Add(speak);
            speak.Click += delegate { Speak(true); };
            Activated += delegate { RefreshSavedState(); };
            RefreshSavedState();
            FormClosed += delegate
            {
                if (speaker == null) return;
                try { speaker.Speak("", 3); } catch { }
                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(speaker); } catch { }
                speaker = null;
            };
        }
        private void SetSavedState(bool saved)
        {
            collect.Text = saved ? "已加入生词本" : "加入生词本";
            collect.Enabled = true;
            note.ReadOnly = saved;
        }
        private void RefreshSavedState()
        {
            try
            {
                Entry saved = findSaved();
                SetSavedState(saved != null);
                if (saved != null) note.Text = saved.Note ?? "";
            }
            catch { } // A storage failure is reported if the user actually tries saving.
        }
        private void Speak(bool showError)
        {
                try
                {
                    if (speaker == null)
                    {
                        speaker = Activator.CreateInstance(Type.GetTypeFromProgID("SAPI.SpVoice", true));
                        dynamic voices = speaker.GetVoices("Language=409");
                        if ((int)voices.Count == 0) voices = speaker.GetVoices("Language=809");
                        if ((int)voices.Count == 0) throw new InvalidOperationException("请先在 Windows 设置中安装英语语音包。");
                        speaker.Voice = voices.Item(0);
                    }
                    speaker.Speak(ReadingInteraction.NormalizeWord(queriedText).Replace('-', ' '), 3); // async + purge previous speech
                }
                catch (Exception ex)
                {
                    if (speaker != null) { try { System.Runtime.InteropServices.Marshal.ReleaseComObject(speaker); } catch { } speaker = null; }
                    if (showError) MessageBox.Show(this, ex.Message, "朗读不可用");
                }
        }
        public void SetDefinition(Definition definition)
        {
            string display = definition.Meaning.Replace("\r\n", "\n").Replace("\n", "\r\n");
            if (meaning.Text != display) meaning.Text = display;
            string ipa = ReadingInteraction.VisiblePhonetic(definition.Phonetic);
            phonetic.Text = ReadingInteraction.IsWord(queriedText) ? queriedText + (ipa.Length > 0 ? "\n" + ipa : "") : "选中文本翻译";
            phonetic.Height = ipa.Length > 0 ? 40 : 22;
            meaning.Location = new Point(10, ipa.Length > 0 ? 47 : 29);
            meaning.Height = ipa.Length > 0 ? 73 : 91;
        }
        public void SetMessage(string text) { meaning.Text = text.Replace("\r\n", "\n").Replace("\n", "\r\n"); }
    }

    public class SettingsForm : Form
    {
        public SettingsForm()
        {
            Text = "接口设置 · 开源版"; ClientSize = new Size(470, 330); StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 10); FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false;
            string[] keys = Storage.LoadKeys();
            Controls.Add(new Label { Text = "本地词典：" + LocalLexicon.Count.ToString("N0") + " 词\n默认离线查词；在线翻译需使用自己的密钥。", Location = new Point(18, 15), Size = new Size(430, 46) });
            Controls.Add(new Label { Text = "有道应用 ID（AppKey）", Location = new Point(18, 66), AutoSize = true });
            var key = new TextBox { Text = keys[0], Location = new Point(18, 93), Width = 430 }; Controls.Add(key);
            Controls.Add(new Label { Text = "有道应用密钥（AppSecret）", Location = new Point(18, 128), AutoSize = true });
            var secret = new TextBox { Text = keys[1], UseSystemPasswordChar = true, Location = new Point(18, 155), Width = 430 }; Controls.Add(secret);
            Controls.Add(new Label { Text = "本地命中不联网；密钥由 Windows 加密保存在本机。\n在线请求仅发送主动选中的待译内容。", Location = new Point(18, 195), Size = new Size(435, 42) });
            var online = new CheckBox { Text = "开启在线翻译（使用自己的有道额度）", Checked = OnlineTranslation.Enabled, Location = new Point(18, 240), AutoSize = true }; Controls.Add(online);
            var save = new Button { Text = "保存", Location = new Point(338, 283), Size = new Size(110, 32) }; Controls.Add(save);
            save.Click += delegate
            {
                try
                {
                    if (online.Checked && (String.IsNullOrWhiteSpace(key.Text) || String.IsNullOrWhiteSpace(secret.Text))) throw new InvalidOperationException("请填写全部接入信息。");
                    Storage.SaveKeys(key.Text.Trim(), secret.Text.Trim()); OnlineTranslation.Enabled = online.Checked; Close();
                }
                catch (Exception ex) { MessageBox.Show(this, ex.Message); }
            };
        }
    }

    public class BookForm : Form
    {
        private DataGridView grid;
        private readonly Func<System.Collections.Generic.List<Entry>> loadEntries;
        public BookForm(Func<System.Collections.Generic.List<Entry>> loadEntries = null, Action<string> onRemoved = null)
        {
            this.loadEntries = loadEntries ?? Storage.Load;
            Text = "生词本"; StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 10);
            AutoScaleDimensions = new SizeF(96, 96); AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(900, 520); MinimumSize = new Size(680, 400);
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 3 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label { Text = "生词本", AutoSize = true, Font = new Font(Font.FontFamily, 14, FontStyle.Bold), Margin = new Padding(0, 0, 0, 12) }, 0, 0);
            grid = new DataGridView { Dock = DockStyle.Fill, Margin = Padding.Empty, AutoGenerateColumns = false, ReadOnly = true, AllowUserToAddRows = false, RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize, AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells };
            grid.ColumnHeadersDefaultCellStyle.Padding = new Padding(8);
            grid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            grid.DefaultCellStyle.Padding = new Padding(6, 4, 6, 4);
            grid.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            foreach (var pair in new[] { new[] { "Word", "单词" }, new[] { "Sentence", "原句" }, new[] { "Note", "个人笔记" }, new[] { "Created", "收藏时间" } })
                grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = pair[0], HeaderText = pair[1], MinimumWidth = 110 });
            grid.Columns[1].FillWeight = 180; grid.Columns[2].FillWeight = 130;
            grid.CellFormatting += delegate(object sender, DataGridViewCellFormattingEventArgs args)
            {
                if (args.ColumnIndex < 0 || grid.Columns[args.ColumnIndex].DataPropertyName != "Created") return;
                DateTimeOffset date;
                if (DateTimeOffset.TryParse(Convert.ToString(args.Value), out date)) { args.Value = date.ToString("yyyy-MM-dd HH:mm"); args.FormattingApplied = true; }
            };
            var bar = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(0, 12, 0, 0), WrapContents = true };
            var refresh = new Button { Text = "刷新", AutoSize = true }; refresh.Click += delegate { Safe(Reload); };
            var remove = new Button { Text = "删除所选", AutoSize = true }; remove.Click += delegate { Safe(delegate
            {
                if (grid.CurrentRow == null) return;
                var entry = grid.CurrentRow.DataBoundItem as Entry;
                if (entry == null) return;
                Storage.Remove(entry.Word);
                if (onRemoved != null) onRemoved(entry.Word);
                Reload();
            }); };
            var export = new Button { Text = "导出 Excel / CSV", AutoSize = true }; export.Click += delegate { Safe(delegate
            {
                using (var dialog = new SaveFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx|CSV 文件 (*.csv)|*.csv", DefaultExt = "xlsx", AddExtension = true, FileName = "生词本" })
                    if (dialog.ShowDialog(this) == DialogResult.OK)
                    {
                        var entries = this.loadEntries();
                        if (String.Equals(Path.GetExtension(dialog.FileName), ".xlsx", StringComparison.OrdinalIgnoreCase)) ExcelExport.Save(dialog.FileName, entries);
                        else if (String.Equals(Path.GetExtension(dialog.FileName), ".csv", StringComparison.OrdinalIgnoreCase))
                        {
                            var lines = entries.Select(e => String.Join(",", new[] { e.Word, e.Sentence, e.Note, e.Created }.Select(Storage.Csv)));
                            File.WriteAllText(dialog.FileName, "单词,原句,个人笔记,收藏时间\r\n" + String.Join("\r\n", lines), new UTF8Encoding(true));
                        }
                        else throw new InvalidOperationException("请选择 .xlsx 或 .csv 文件格式。");
                    }
            }); };
            foreach (var button in new[] { refresh, remove, export }) { button.Padding = new Padding(10, 4, 10, 4); button.Margin = new Padding(0, 0, 12, 0); }
            bar.Controls.AddRange(new Control[] { refresh, remove, export });
            layout.Controls.Add(grid, 0, 1); layout.Controls.Add(bar, 0, 2); Controls.Add(layout); Reload();
        }
        private void Reload() { grid.DataSource = loadEntries().OrderByDescending(e => e.Created).ToList(); }
        private void Safe(Action action) { try { action(); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "操作失败"); } }
    }
}
