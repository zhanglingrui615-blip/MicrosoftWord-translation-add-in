using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

[assembly: ComVisible(false)]
[assembly: System.Reflection.AssemblyVersion("0.1.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("0.4.1.0")]
namespace WordVocabulary
{
    [ComImport, ComVisible(true), Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object app, [In] int mode, [In, MarshalAs(UnmanagedType.IDispatch)] object instance, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection([In] int mode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }
    [ComImport, ComVisible(true), Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility { [DispId(1)] [return: MarshalAs(UnmanagedType.BStr)] string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)] string ribbonId); }

    [ComVisible(true), Guid("757C1B15-D248-4DC4-827B-BFD9D12B49A1"), ProgId("WordVocabulary.Connect"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        public Connect() { Log("Constructor 0.4.1 silent vocabulary toggle and wave cleanup"); }
        private static void Log(string text)
        {
            try { System.IO.File.AppendAllText(Storage.PathFor("startup.log"), DateTime.Now.ToString("o") + " " + text + Environment.NewLine); } catch { }
        }
        private dynamic app;
        private System.Windows.Forms.Timer timer;
        private bool enabled, wasDown;
        private Point downPoint;
        private Point releasePoint;
        private DateTime translateAfter;
        private bool waitForSelection;
        private bool dragged;
        private bool downInText;
        private WordCard card;
        private CancellationTokenSource pending;
        private int generation;
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(Point point);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder text, int count);
        public void OnConnection(object application, int mode, object instance, ref Array custom)
        {
            Log("OnConnection begin");
            try {
            app = application;
            // Word callbacks and this timer run on Word's UI STA thread.
            SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            timer = new System.Windows.Forms.Timer { Interval = 40 };
            timer.Tick += Tick;
            timer.Start();
            Log("OnConnection complete");
            } catch (Exception ex) { Log(ex.ToString()); throw; }
        }
        public string GetCustomUI(string ribbonId)
        {
            Log("GetCustomUI " + ribbonId);
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon><tabs><tab id='vocabTab' label='翻译助手'>
    <group id='translationGroup' getLabel='GetEmptyGroupLabel'>
      <box id='translationColumn' boxStyle='vertical'>
      <labelControl id='translationTopSpace' label='&#160;'/>
      <toggleButton id='reading' label='翻译模式' size='normal' showLabel='true' onAction='ToggleReading'
        screentip='开启或关闭翻译模式' supertip='开启后，点击单词直接查词并朗读；拖动选中短语或句子，松开鼠标即翻译选中内容。点击空白处关闭词条，再次点击此按钮退出翻译模式。'/>
      <labelControl id='translationBottomSpace' label='&#160;'/>
      </box>
    </group>
    <group id='vocabularyGroup' getLabel='GetEmptyGroupLabel'>
      <box id='vocabularyColumn' boxStyle='vertical'>
      <labelControl id='vocabularyTopSpace' label='&#160;'/>
      <button id='book' label='生词本' size='normal' showLabel='true' onAction='ShowBook'
        screentip='打开生词本' supertip='查看、删除或导出生词及原句。'/>
      <labelControl id='vocabularyBottomSpace' label='&#160;'/>
      </box>
    </group>
    <group id='preferencesGroup' getLabel='GetEmptyGroupLabel'>
      <box id='preferencesColumn' boxStyle='vertical'>
      <labelControl id='preferencesTopSpace' label='&#160;'/>
      <button id='settings' label='接口设置' size='normal' showLabel='true' onAction='ShowSettings'
        screentip='设置有道翻译' supertip='查看本地词典数量，配置未收录单词和句子翻译使用的有道密钥。'/>
      <labelControl id='preferencesBottomSpace' label='&#160;'/>
      </box>
    </group>
  </tab></tabs></ribbon>
</customUI>";
        }
        public string GetEmptyGroupLabel(object control) { return String.Empty; }
        public void ToggleReading(object control, bool pressed) { enabled = pressed; wasDown = false; waitForSelection = false; CloseCard(); }
        public void ShowSettings(object control)
        {
            try { using (var form = new SettingsForm()) form.ShowDialog(); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "设置未能打开"); }
        }
        public void ShowBook(object control)
        {
            try { new BookForm(null, delegate(string word) { WordMarks.SyncWord((object)app, word, false); }).Show(); } catch (Exception ex) { MessageBox.Show(ex.Message, "生词本未能打开"); }
        }
        public void LookupSelection(object control)
        {
            try
            {
                dynamic range = app.Selection.Range.Duplicate;
                bool selected = (int)range.Start != (int)range.End;
                if (!selected)
                {
                    int at = (int)range.Start;
                    dynamic paragraph = range.Paragraphs[1].Range.Duplicate;
                    paragraph.SetRange(Math.Max((int)paragraph.Start, at - 128), Math.Min((int)paragraph.End, at + 128));
                    Match token = ReadingInteraction.WordNearCaret((string)paragraph.Text, at - (int)paragraph.Start);
                    if (token.Success) { int start = (int)paragraph.Start + token.Index; range.SetRange(start, start + token.Length); }
                    else range.Expand(2);
                }
                ShowRange(range, Cursor.Position, selected);
            }
            catch (Exception ex) { MessageBox.Show("请先打开文档，选中单词或句子。\n" + ex.Message); }
        }
        private void Tick(object sender, EventArgs e)
        {
            if (!enabled) return;
            bool down = (GetAsyncKeyState(1) & 0x8000) != 0;
            if (down && !wasDown)
            {
                downPoint = Cursor.Position; waitForSelection = false;
                var targetClass = new StringBuilder(64);
                GetClassName(WindowFromPoint(downPoint), targetClass, 64);
                downInText = targetClass.ToString() == "_WwG";
                if (downInText) CloseCard();
            }
            bool released = wasDown && !down;
            wasDown = down;
            if (released)
            {
                if (!downInText) { waitForSelection = false; return; }
                releasePoint = Cursor.Position;
                dragged = Math.Abs(releasePoint.X - downPoint.X) > 5 || Math.Abs(releasePoint.Y - downPoint.Y) > 5;
                translateAfter = DateTime.UtcNow.AddMilliseconds(120);
                waitForSelection = true;
            }
            if (!waitForSelection || down || DateTime.UtcNow < translateAfter) return;
            waitForSelection = false;
            var name = new StringBuilder(64);
            GetClassName(GetForegroundWindow(), name, 64);
            if (name.ToString() != "OpusApp") return;
            Point point = releasePoint;
            try
            {
                if ((int)app.Documents.Count == 0) return;
                dynamic range;
                if (dragged)
                {
                    if (!downInText) return;
                    range = app.Selection.Range.Duplicate;
                    if ((int)range.Start == (int)range.End || String.IsNullOrWhiteSpace((string)range.Text)) { DismissSelection(); return; }
                    ShowRange(range, point, true);
                }
                else
                {
                    dynamic hit = app.ActiveWindow.RangeFromPoint(point.X, point.Y);
                    if (hit == null) { DismissSelection(); return; }
                    object clickedRange;
                    if (!TryGetClickedWord((object)hit, point, out clickedRange)) { DismissSelection(); return; }
                    range = clickedRange;
                    ShowRange(range, point, false);
                }
            }
            catch { CloseCard(); } // Shapes, dialogs, margins, and busy Word are not text targets.
        }
        private bool TryGetClickedWord(object rawHit, Point point, out object result)
        {
            result = null;
            dynamic hit = rawHit;
            // RangeFromPoint can return the nearest caret, including in whitespace.
            // Try both sides of that caret, but accept only a real non-space character under the pointer.
            int position;
            try { position = (int)hit.Start; } catch { return false; }
            for (int candidate = 0; candidate < 2; candidate++)
            {
                try
                {
                    dynamic range = hit.Paragraphs[1].Range.Duplicate;
                    int paragraphStart = (int)range.Start;
                    int paragraphEnd = (int)range.End;
                    int at = Math.Max(paragraphStart, position - candidate);
                    int contextStart = Math.Max(paragraphStart, at - 128);
                    range.SetRange(contextStart, Math.Min(paragraphEnd, at + 128));
                    string text = (string)range.Text;
                    Match match = ReadingInteraction.WordNearCaret(text, at - contextStart);
                    if (!match.Success || match.Length > 96) continue;
                    int start = contextStart + match.Index;
                    range.SetRange(start, start + match.Length);
                    for (int i = 1; i <= (int)range.Characters.Count; i++)
                    {
                        dynamic character = range.Characters[i];
                        int left = 0, top = 0, width = 0, height = 0;
                        app.ActiveWindow.GetPoint(out left, out top, out width, out height, character);
                        if (ReadingInteraction.IsGlyphHit(point, new Rectangle(left, top, width, height), (string)character.Text))
                        { result = range; return true; }
                    }
                }
                catch { }
            }
            return false;
        }
        private void DismissSelection()
        {
            CloseCard();
            try { app.Selection.Collapse(0); } catch { }
        }
        private async void ShowRange(dynamic range, Point point, bool selected)
        {
            try { await ShowRangeAsync((object)range, point, selected); }
            catch (Exception ex) { CloseCard(); MessageBox.Show("无法读取当前选区，请重新选择。\n" + ex.Message, "选中翻译"); }
        }
        private async System.Threading.Tasks.Task ShowRangeAsync(object rawRange, Point point, bool selected)
        {
            dynamic range = rawRange;
            string text = (string)range.Text;
            string word;
            int offset;
            if (selected)
            {
                word = text.Trim(); offset = text.Length - text.TrimStart().Length;
                if (String.IsNullOrWhiteSpace(word)) { CloseCard(); return; }
                if (word.Length > ReadingInteraction.MaxSelectionLength) { CloseCard(); MessageBox.Show("一次最多翻译 6000 个字符，请缩小选中范围。"); return; }
            }
            else
            {
                Match match = Regex.Match(text, @"^\s*(" + ReadingInteraction.WordPattern + @")[\s.,;:!?]*$");
                if (!match.Success || match.Groups[1].Length > 96) { CloseCard(); return; }
                word = match.Groups[1].Value; offset = match.Groups[1].Index;
            }
            int start = (int)range.Start + offset;
            range.SetRange(start, start + word.Length);
            try { WordMarks.ApplyRange((object)range, Storage.FindEntry(Storage.Load(), word) != null); } catch { }
            range.Select(); // Keep the actual queried text selected in Word.
            Rectangle anchor = new Rectangle(point.X, point.Y, 1, 15);
            try
            {
                int left = 0, top = 0, width = 0, height = 0;
                app.ActiveWindow.GetPoint(out left, out top, out width, out height, range);
                anchor = new Rectangle(left, top, width, height);
            }
            catch
            {
                // Multi-page selections may not be entirely visible: anchor to their visible end.
                try
                {
                    dynamic edge = range.Duplicate;
                    edge.SetRange(Math.Max((int)range.Start, (int)range.End - 1), (int)range.End);
                    int left = 0, top = 0, width = 0, height = 0;
                    app.ActiveWindow.GetPoint(out left, out top, out width, out height, edge);
                    anchor = new Rectangle(left, top, width, height);
                }
                catch { }
            }
            string sentence = "";
            try { sentence = ((string)range.Sentences[1].Text).Trim(); } catch { }
            CloseCard();
            int request = generation;
            pending = new CancellationTokenSource();
            CancellationToken token = pending.Token;
            dynamic captured = range;
            WordCard current = new WordCard(word, sentence, delegate(string note)
            {
                // Word ranges track edits. Verify the word before applying formatting.
                if (!String.Equals(((string)captured.Text).Trim(), word, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("原文已变化，请重新选中后收藏。");
                Storage.Add(word, sentence, note);
                WordMarks.ApplyRange((object)captured, true);
                WordMarks.SyncWord((object)app, word, true);
            }, remove: delegate
            {
                Storage.Remove(word);
                try { if (String.Equals(((string)captured.Text).Trim(), word, StringComparison.OrdinalIgnoreCase)) WordMarks.ApplyRange((object)captured, false); } catch { }
                WordMarks.SyncWord((object)app, word, false);
            });
            card = current;
            current.FormClosed += delegate { if (card == current) CloseCard(); };
            Rectangle screen = Screen.FromPoint(new Point(anchor.Left, anchor.Bottom)).WorkingArea;
            current.Location = ReadingInteraction.PlaceCard(anchor, current.Size, screen);
            current.Show(new WordWindow(new IntPtr((int)app.ActiveWindow.Hwnd)));
            bool complete = false;
            var progress = new Progress<Definition>(delegate(Definition partial)
            {
                if (!complete && generation == request && !current.IsDisposed) current.SetDefinition(partial);
            });
            try
            {
                Definition result = await DictionaryProvider.Lookup(word, sentence, token, progress);
                complete = true;
                if (generation == request && !current.IsDisposed) current.SetDefinition(result);
            }
            catch (OperationCanceledException) { complete = true; if (generation == request && !current.IsDisposed) current.SetMessage("查询超时，请重新点击单词。"); }
            catch (Exception ex) { complete = true; if (generation == request && !current.IsDisposed) current.SetMessage(ex.Message); }
        }
        private sealed class WordWindow : IWin32Window
        {
            public IntPtr Handle { get; private set; }
            public WordWindow(IntPtr handle) { Handle = handle; }
        }
        private void CloseCard()
        {
            generation++;
            if (pending != null) { pending.Cancel(); pending.Dispose(); pending = null; }
            WordCard old = card; card = null;
            if (old != null && !old.IsDisposed) old.Close();
        }
        private void Stop() { enabled = false; if (timer != null) { timer.Stop(); timer.Dispose(); timer = null; } CloseCard(); app = null; }
        public void OnDisconnection(int mode, ref Array custom) { Stop(); }
        public void OnBeginShutdown(ref Array custom) { Stop(); }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
    }
}
