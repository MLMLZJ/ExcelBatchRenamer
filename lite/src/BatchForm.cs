using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelBatchRename
{
    public class BatchForm : Form
    {
        private readonly UiBridge _ui = new UiBridge();
        private readonly Timer _timer = new Timer();

        private TextBox _txtFolder;
        private TextBox _txtFilter;
        private TextBox _txtPrefix;
        private TextBox _txtSuffix;
        private TextBox _txtFind;
        private TextBox _txtReplace;
        private CheckBox _chkRegex;
        private ComboBox _cbCase;
        private RadioButton _rbOpRename;
        private RadioButton _rbOpCopy;
        private TextBox _txtOutput;
        private Button _btnOutput;
        private RadioButton _rbDupSkip;
        private RadioButton _rbDupAuto;
        private Button _btnPreview;
        private Button _btnRun;
        private Button _btnUndo;
        private Button _btnClear;

        private PreviewResult _lastPreview;
        private BatchRenameOptions _lastOptions;
        private bool _busy;

        public BatchForm()
        {
            Text = "普通批量改名";
            StartPosition = FormStartPosition.CenterParent;
            Font = UiKit.UiFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(980, 680);
            MinimumSize = new Size(820, 560);
            Icon icon = UiKit.AppIcon();
            if (icon != null) Icon = icon;

            BuildUi();
            _timer.Interval = 100;
            _timer.Tick += delegate { _ui.Flush(); };
            _timer.Start();
        }

        private void BuildUi()
        {
            TableLayoutPanel root = UiKit.Root();
            root.AutoScroll = true;
            Controls.Add(root);

            _txtFolder = UiKit.Input(520);
            _txtFilter = UiKit.Input(200);
            _txtFilter.Text = "*.pdf";

            FlowLayoutPanel top = UiKit.Column(
                UiKit.Row(UiKit.Text("1. 选择文件夹：", 0), _txtFolder,
                          UiKit.Button("浏览…", delegate { BrowseFolder(); })),
                UiKit.Row(UiKit.Text("2. 文件类型：", 0), _txtFilter,
                          UiKit.Hint("如 *.pdf   *.docx   *.*（多个用空格分隔）")));
            root.Controls.Add(top, 0, 0);

            _txtPrefix = UiKit.Input(150);
            _txtSuffix = UiKit.Input(150);
            _txtFind = UiKit.Input(150);
            _txtReplace = UiKit.Input(150);
            _chkRegex = UiKit.Check("正则表达式", false);
            _cbCase = new ComboBox();
            _cbCase.DropDownStyle = ComboBoxStyle.DropDownList;
            _cbCase.Width = 120;
            _cbCase.Margin = new Padding(0, 4, 6, 4);
            _cbCase.Items.AddRange(new object[] { "不变", "全部大写", "全部小写", "首字母大写" });
            _cbCase.SelectedIndex = 0;

            _rbOpRename = UiKit.Radio("直接改名", true, delegate { ToggleOutput(); });
            _rbOpCopy = UiKit.Radio("复制到新文件夹（原文件保留，推荐）", false, delegate { ToggleOutput(); });
            _txtOutput = UiKit.Input(420);
            _txtOutput.Enabled = false;
            _btnOutput = UiKit.Button("浏览…", delegate { BrowseOutput(); });
            _btnOutput.Enabled = false;

            _rbDupSkip = UiKit.Radio("跳过", true, null);
            _rbDupAuto = UiKit.Radio("自动加序号", false, null);

            FlowLayoutPanel rules = UiKit.Column(
                UiKit.Row(UiKit.Text("添加前缀：", 0), _txtPrefix,
                          UiKit.Text("    添加后缀：", 0), _txtSuffix,
                          UiKit.Hint("加在扩展名前")),
                UiKit.Row(UiKit.Text("查找：", 0), _txtFind,
                          UiKit.Text("    替换为：", 0), _txtReplace, _chkRegex),
                UiKit.Row(UiKit.Text("大小写转换：", 0), _cbCase,
                          UiKit.Hint("执行顺序：查找替换 → 大小写转换 → 加前缀 → 加后缀")),
                UiKit.Row(UiKit.Text("操作模式：", 0), _rbOpRename, _rbOpCopy),
                UiKit.Row(UiKit.Text("输出文件夹：", 0), _txtOutput, _btnOutput),
                UiKit.Row(UiKit.Text("重名处理：", 0), _rbDupSkip, _rbDupAuto));
            root.Controls.Add(UiKit.Section("改名规则", rules), 0, 1);

            _btnPreview = UiKit.Button("预览（不改文件）", delegate { DoPreview(); });
            _btnRun = UiKit.Button("执行改名", delegate { DoRun(); });
            _btnUndo = UiKit.Button("撤销", delegate { DoUndo(); });
            _btnClear = UiKit.Button("清空日志", delegate { _ui.Clear(); });
            FlowLayoutPanel btns = UiKit.Row(_btnPreview, _btnRun, _btnUndo, _btnClear);
            btns.Dock = DockStyle.Fill;
            root.Controls.Add(btns, 0, 2);

            root.Controls.Add(BuildLogPanel(), 0, 3);
        }

        private Control BuildLogPanel()
        {
            Panel logPanel = new Panel();
            logPanel.Dock = DockStyle.Fill;
            logPanel.Margin = new Padding(0);

            ProgressBar bar = new ProgressBar();
            bar.Dock = DockStyle.Top;
            bar.Height = 14;
            bar.Minimum = 0;
            bar.Maximum = 1;

            Label logLabel = new Label();
            logLabel.Text = "运行日志：";
            logLabel.Dock = DockStyle.Top;
            logLabel.Height = 24;
            logLabel.TextAlign = ContentAlignment.MiddleLeft;

            RichTextBox box = new RichTextBox();
            box.Dock = DockStyle.Fill;
            box.ReadOnly = true;
            box.BackColor = Color.White;
            box.BorderStyle = BorderStyle.FixedSingle;
            box.Font = UiKit.LogFont;
            box.HideSelection = false;
            box.DetectUrls = false;

            logPanel.Controls.Add(box);
            logPanel.Controls.Add(logLabel);
            logPanel.Controls.Add(bar);
            logPanel.Controls.SetChildIndex(box, 0);
            logPanel.Controls.SetChildIndex(logLabel, 1);
            logPanel.Controls.SetChildIndex(bar, 2);

            _ui.Box = box;
            _ui.Bar = bar;
            return logPanel;
        }

        private void BrowseFolder()
        {
            using (FolderBrowserDialog d = new FolderBrowserDialog())
            {
                d.Description = "选择要批量改名的文件夹";
                if (!string.IsNullOrEmpty(_txtFolder.Text) && Directory.Exists(_txtFolder.Text))
                    d.SelectedPath = _txtFolder.Text;
                if (d.ShowDialog(this) == DialogResult.OK) _txtFolder.Text = d.SelectedPath;
            }
        }

        private void BrowseOutput()
        {
            using (FolderBrowserDialog d = new FolderBrowserDialog())
            {
                d.Description = "选择输出文件夹";
                if (d.ShowDialog(this) == DialogResult.OK) _txtOutput.Text = d.SelectedPath;
            }
        }

        private void ToggleOutput()
        {
            bool copy = _rbOpCopy.Checked;
            _txtOutput.Enabled = copy;
            _btnOutput.Enabled = copy;
        }

        private bool CollectOptions(out BatchRenameOptions opt)
        {
            opt = null;
            string folder = _txtFolder.Text.Trim();
            if (!Directory.Exists(folder))
            {
                MessageBox.Show(this, "请先选择文件夹", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            BatchRenameOptions o = new BatchRenameOptions();
            o.Folder = folder;
            o.FileFilter = _txtFilter.Text;
            o.AddPrefix = _txtPrefix.Text;
            o.AddSuffix = _txtSuffix.Text;
            o.FindText = _txtFind.Text;
            o.ReplaceText = _txtReplace.Text;
            o.UseRegex = _chkRegex.Checked;
            o.CaseMode = _cbCase.SelectedItem == null ? "不变" : _cbCase.SelectedItem.ToString();
            o.OpMode = _rbOpCopy.Checked ? "copy" : "rename";
            o.OutputFolder = _txtOutput.Text.Trim();
            o.DupMode = _rbDupAuto.Checked ? "auto" : "skip";
            if (o.OpMode == "copy")
            {
                if (!Directory.Exists(o.OutputFolder))
                {
                    MessageBox.Show(this, "复制模式下请先选择输出文件夹", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                if (RenameEngine.SamePath(o.OutputFolder, o.Folder))
                {
                    MessageBox.Show(this, "输出文件夹不能和原文件夹相同", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }
            opt = o;
            return true;
        }

        private void DoPreview()
        {
            if (_busy) return;
            BatchRenameOptions opt;
            if (!CollectOptions(out opt)) return;

            _ui.Log("");
            _ui.Log("======= 预览模式，不会修改文件 =======", "info");
            _ui.ResetProgress();
            SetBusy(true);
            Task.Run(delegate
            {
                return RenameEngine.PreviewBatch(opt,
                    delegate(LogItem li) { _ui.Log(li); },
                    delegate(int cur, int max) { _ui.Progress(cur, max); });
            }).ContinueWith(delegate(Task<PreviewResult> t)
            {
                SetBusy(false);
                _ui.ResetProgress();
                if (t.IsFaulted)
                {
                    Exception ex = t.Exception == null ? null : t.Exception.GetBaseException();
                    string m = ex == null ? "未知错误" : ex.Message;
                    _ui.Log("预览失败：" + m, "miss");
                    MessageBox.Show(this, m, "预览失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                _lastPreview = t.Result;
                _lastOptions = opt;
                if (_lastPreview.Items.Count == 0)
                    MessageBox.Show(this, "没有可处理的文件，详见日志。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    _ui.Log("确认无误后点击【执行改名】", "info");
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void DoRun()
        {
            if (_busy) return;
            if (_lastPreview == null || _lastPreview.Items.Count == 0)
            {
                MessageBox.Show(this, "请先执行【预览】", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            BatchRenameOptions opt;
            if (!CollectOptions(out opt)) return;
            if (_lastOptions == null || !SameOptions(opt, _lastOptions))
            {
                MessageBox.Show(this, "规则或路径已修改，请重新点击【预览】再执行。", "提示",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool isCopy = opt.OpMode == "copy";
            string warn = _lastPreview.DupCount > 0
                ? string.Format("\n\n注意：预览中发现 {0} 个重名跳过。这些文件不会被处理，是否仍要继续？", _lastPreview.DupCount)
                : "";
            string msg = isCopy
                ? "确认复制改名！原文件保留，副本输出到：\n" + opt.OutputFolder + warn + "\n\n确认继续？"
                : "确认执行改名！将直接修改原文件名！\n建议先备份！" + warn + "\n\n确认继续？";
            if (MessageBox.Show(this, msg, "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

            _ui.Log("");
            _ui.Log("======= 执行开始 =======", "info");
            _ui.ResetProgress();
            SetBusy(true);
            Task.Run(delegate
            {
                return RenameEngine.ExecuteBatch(_lastPreview.Items, opt,
                    delegate(LogItem li) { _ui.Log(li); },
                    delegate(int cur, int max) { _ui.Progress(cur, max); });
            }).ContinueWith(delegate(Task<ExecuteResult> t)
            {
                SetBusy(false);
                _ui.ResetProgress();
                if (t.IsFaulted)
                {
                    Exception ex = t.Exception == null ? null : t.Exception.GetBaseException();
                    string m = ex == null ? "未知错误" : ex.Message;
                    _ui.Log("执行失败：" + m, "miss");
                    MessageBox.Show(this, m, "执行失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                ExecuteResult r = t.Result;
                _ui.Log(string.Format("完成：成功 {0} 个，失败 {1} 个", r.Success, r.Failed), "info");
                MessageBox.Show(this, string.Format("执行完成：成功 {0} 个，失败 {1} 个。", r.Success, r.Failed),
                                "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private static bool SameOptions(BatchRenameOptions a, BatchRenameOptions b)
        {
            return a.Folder == b.Folder && a.FileFilter == b.FileFilter && a.AddPrefix == b.AddPrefix
                && a.AddSuffix == b.AddSuffix && a.FindText == b.FindText && a.ReplaceText == b.ReplaceText
                && a.UseRegex == b.UseRegex && a.CaseMode == b.CaseMode && a.OpMode == b.OpMode
                && a.OutputFolder == b.OutputFolder && a.DupMode == b.DupMode;
        }

        private void DoUndo()
        {
            if (_busy) return;
            string folder = _txtFolder.Text.Trim();
            if (!Directory.Exists(folder))
            {
                MessageBox.Show(this, "请先选择文件夹（撤销时需要）", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string fp;
            using (OpenFileDialog d = new OpenFileDialog())
            {
                d.Title = "选择记录文件";
                d.Filter = "CSV 记录文件 (*.csv)|*.csv|所有文件 (*.*)|*.*";
                d.InitialDirectory = folder;
                if (d.ShowDialog(this) != DialogResult.OK) return;
                fp = d.FileName;
            }
            bool isCopy = RenameEngine.IsCopyModeRecord(fp, folder);
            string recordDir = Path.GetDirectoryName(Path.GetFullPath(fp));
            string msg = isCopy
                ? "检测为复制模式，将删除输出文件夹中的副本：\n" + recordDir + "\n\n确认继续？"
                : "将根据记录把文件名改回原名。\n\n确认继续？";
            if (MessageBox.Show(this, msg, "确认撤销", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

            _ui.Log("");
            _ui.Log("======= 撤销开始 =======", "info");
            _ui.ResetProgress();
            SetBusy(true);
            Task.Run(delegate
            {
                return RenameEngine.Undo(fp, folder,
                    delegate(LogItem li) { _ui.Log(li); },
                    delegate(int cur, int max) { _ui.Progress(cur, max); });
            }).ContinueWith(delegate(Task<ExecuteResult> t)
            {
                SetBusy(false);
                _ui.ResetProgress();
                if (t.IsFaulted)
                {
                    Exception ex = t.Exception == null ? null : t.Exception.GetBaseException();
                    string m = ex == null ? "未知错误" : ex.Message;
                    _ui.Log("撤销失败：" + m, "miss");
                    MessageBox.Show(this, m, "撤销失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                ExecuteResult r = t.Result;
                _ui.Log(string.Format("撤销完成：成功 {0} 个，失败 {1} 个", r.Success, r.Failed), "info");
                MessageBox.Show(this, string.Format("撤销完成：成功 {0} 个，失败 {1} 个。", r.Success, r.Failed),
                                "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void SetBusy(bool busy)
        {
            _busy = busy;
            UseWaitCursor = busy;
            _btnPreview.Enabled = !busy;
            _btnRun.Enabled = !busy;
            _btnUndo.Enabled = !busy;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _timer.Stop();
            _timer.Dispose();
            base.OnFormClosing(e);
        }
    }
}