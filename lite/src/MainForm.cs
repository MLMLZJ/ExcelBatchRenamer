using System;
using System.Collections.Concurrent;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelBatchRename
{
    internal static class Colors
    {
        public static readonly Color Miss = Color.FromArgb(0xD3, 0x2F, 0x2F);
        public static readonly Color Dup = Color.FromArgb(0xF5, 0x7C, 0x00);
        public static readonly Color Ok = Color.FromArgb(0x2E, 0x7D, 0x32);
        public static readonly Color Info = Color.FromArgb(0x15, 0x65, 0xC0);
        public static readonly Color Normal = Color.FromArgb(0x33, 0x33, 0x33);
    }

    /// <summary>UI 日志与进度：后台线程只入队，界面计时器负责刷新</summary>
    internal class UiBridge
    {
        private readonly ConcurrentQueue<LogItem> _queue = new ConcurrentQueue<LogItem>();
        private int _progCur;
        private int _progMax;

        public RichTextBox Box;
        public ProgressBar Bar;

        public void Log(string text, string tag = null)
        {
            _queue.Enqueue(new LogItem(text, tag));
        }

        public void Log(LogItem item)
        {
            if (item != null) _queue.Enqueue(item);
        }

        public void Progress(int cur, int max)
        {
            _progCur = cur;
            _progMax = max;
        }

        public void ResetProgress()
        {
            _progCur = 0;
            _progMax = 0;
        }

        public void Flush()
        {
            if (Box == null || Box.IsDisposed) return;
            LogItem item;
            int guard = 0;
            bool changed = false;
            while (_queue.TryDequeue(out item) && guard < 500)
            {
                guard++;
                Box.SelectionStart = Box.TextLength;
                Box.SelectionLength = 0;
                Box.SelectionColor = TagColor(item.Tag);
                Box.AppendText(item.Text + Environment.NewLine);
                changed = true;
            }
            if (changed)
            {
                Box.SelectionColor = Colors.Normal;
                Box.SelectionStart = Box.TextLength;
                Box.ScrollToCaret();
            }
            if (Bar != null && !Bar.IsDisposed)
            {
                int max = _progMax < 1 ? 1 : _progMax;
                if (Bar.Maximum != max) Bar.Maximum = max;
                int v = _progCur < 0 ? 0 : (_progCur > max ? max : _progCur);
                if (Bar.Value != v) Bar.Value = v;
            }
        }

        private static Color TagColor(string tag)
        {
            if (tag == "miss") return Colors.Miss;
            if (tag == "dup") return Colors.Dup;
            if (tag == "ok") return Colors.Ok;
            if (tag == "info") return Colors.Info;
            return Colors.Normal;
        }

        public string AllText()
        {
            return Box == null ? "" : Box.Text;
        }

        public void Clear()
        {
            LogItem ignored;
            while (_queue.TryDequeue(out ignored)) { }
            if (Box != null) Box.Clear();
            ResetProgress();
        }
    }

    public class MainForm : Form
    {
        private readonly UiBridge _ui = new UiBridge();
        private readonly Timer _timer = new Timer();

        private TextBox _txtExcel;
        private TextBox _txtFolder;
        private ComboBox _cbPrefix;
        private ComboBox _cbOld;
        private TextBox _txtStartRow;
        private TextBox _txtEndRow;

        private TextBox _txtSeparator;
        private TextBox _txtExtension;
        private NumericUpDown _numPad;
        private RadioButton _rbDupSkip;
        private RadioButton _rbDupAuto;
        private CheckBox _chkClean;
        private RadioButton _rbOpRename;
        private RadioButton _rbOpCopy;
        private TextBox _txtOutput;
        private Button _btnOutput;

        private Button _btnLoad;
        private Button _btnPreview;
        private Button _btnRename;
        private Button _btnUndo;
        private Button _btnBatch;
        private Button _btnExport;
        private Button _btnClear;

        private ExcelTable _table;
        private PreviewResult _lastPreview;
        private ExcelRenameOptions _lastOptions;
        private bool _busy;

        public MainForm()
        {
            Text = "Excel 批量文件重命名工具 v3.0";
            StartPosition = FormStartPosition.CenterScreen;
            Font = UiKit.UiFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(960, 780);
            MinimumSize = new Size(900, 620);
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

            // ---------- 顶部：文件 / 文件夹 / 列 / 行范围 ----------
            _txtExcel = UiKit.Input(520);
            _txtFolder = UiKit.Input(520);
            _cbPrefix = new ComboBox();
            _cbPrefix.DropDownStyle = ComboBoxStyle.DropDownList;
            _cbPrefix.Width = 260;
            _cbPrefix.Margin = new Padding(0, 4, 6, 4);
            _cbOld = new ComboBox();
            _cbOld.DropDownStyle = ComboBoxStyle.DropDownList;
            _cbOld.Width = 260;
            _cbOld.Margin = new Padding(0, 4, 6, 4);
            _txtStartRow = UiKit.Input(56);
            _txtStartRow.Text = "2";
            _txtEndRow = UiKit.Input(56);

            FlowLayoutPanel top = UiKit.Column(
                UiKit.Row(UiKit.Text("1. 选择 Excel 文件（.xlsx）：", 0), _txtExcel,
                          UiKit.Button("浏览…", delegate { BrowseExcel(); })),
                UiKit.Row(UiKit.Text("2. 选择文件所在文件夹：", 0), _txtFolder,
                          UiKit.Button("浏览…", delegate { BrowseFolder(); })),
                UiKit.Row(UiKit.Text("3. 前缀列（Excel 列名）：", 0), _cbPrefix,
                          UiKit.Hint("取这一列作为新文件名的前缀")),
                UiKit.Row(UiKit.Text("4. 旧名字列（Excel 列名）：", 0), _cbOld,
                          UiKit.Hint("对应磁盘上的原文件名，可带或不带扩展名")),
                UiKit.Row(UiKit.Text("5. 处理行范围：", 0),
                          UiKit.Text("从第", 0), _txtStartRow, UiKit.Text("行  到第", 0), _txtEndRow,
                          UiKit.Text("行", 0),
                          UiKit.Hint("留空=全部；行号含表头行，数据从第 2 行开始")));
            root.Controls.Add(top, 0, 0);

            // ---------- 命名规则 ----------
            _txtSeparator = UiKit.Input(70);
            _txtSeparator.Text = "_";
            _txtExtension = UiKit.Input(70);
            _txtExtension.Text = "pdf";
            _numPad = new NumericUpDown();
            _numPad.Minimum = 0;
            _numPad.Maximum = 10;
            _numPad.Value = 0;
            _numPad.Width = 60;
            _numPad.Margin = new Padding(0, 4, 6, 4);
            _rbDupSkip = UiKit.Radio("跳过", true, null);
            _rbDupAuto = UiKit.Radio("自动加序号", false, null);
            _chkClean = UiKit.Check("自动清洗文件名（去首尾空格、替换 \\ / : * ? \" < > | 等非法字符）", true);
            _rbOpRename = UiKit.Radio("直接改名（原文件被改）", true, delegate { ToggleOutput(); });
            _rbOpCopy = UiKit.Radio("复制到新文件夹（原文件保留，推荐）", false, delegate { ToggleOutput(); });
            _txtOutput = UiKit.Input(420);
            _txtOutput.Enabled = false;
            _btnOutput = UiKit.Button("浏览…", delegate { BrowseOutput(); });
            _btnOutput.Enabled = false;

            FlowLayoutPanel rules = UiKit.Column(
                UiKit.Row(UiKit.Text("连接符：", 0), _txtSeparator,
                          UiKit.Text("    扩展名：", 0), _txtExtension,
                          UiKit.Hint("    支持中文连接符（如 _  -  第  号_），扩展名不带点")),
                UiKit.Row(UiKit.Text("前缀补零：", 0), _numPad,
                          UiKit.Text("    重名处理：", 0), _rbDupSkip, _rbDupAuto,
                          UiKit.Hint("    0=不补零；填 3 表示 1 → 001")),
                UiKit.Row(_chkClean),
                UiKit.Row(UiKit.Text("操作模式：", 0), _rbOpRename, _rbOpCopy),
                UiKit.Row(UiKit.Text("输出文件夹：", 0), _txtOutput, _btnOutput));
            root.Controls.Add(UiKit.Section("命名规则", rules), 0, 1);

            // ---------- 按钮 ----------
            _btnLoad = UiKit.Button("加载 Excel 列", delegate { LoadExcel(); });
            _btnPreview = UiKit.Button("预览改名（不改文件）", delegate { DoPreview(); });
            _btnRename = UiKit.Button("执行正式改名", delegate { DoRename(); });
            _btnUndo = UiKit.Button("撤销改名", delegate { DoUndo(); });
            _btnBatch = UiKit.Button("普通批量改名", delegate { new BatchForm().Show(this); });
            _btnExport = UiKit.Button("导出日志", delegate { ExportLog(); });
            _btnClear = UiKit.Button("清空日志", delegate { _ui.Clear(); });
            FlowLayoutPanel btns = UiKit.Column(
                UiKit.Row(_btnLoad, _btnPreview, _btnRename, _btnUndo),
                UiKit.Row(_btnBatch, _btnExport, _btnClear));
            btns.Margin = new Padding(0, 6, 0, 4);
            root.Controls.Add(btns, 0, 2);

            // ---------- 进度 + 日志 ----------
            root.Controls.Add(BuildLogPanel(), 0, 3);
        }

        private Control BuildLogPanel()
        {
            Panel logPanel = new Panel();
            logPanel.Dock = DockStyle.Fill;
            logPanel.Margin = new Padding(0);
            logPanel.Padding = new Padding(0, 2, 0, 0);

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
            logLabel.Margin = new Padding(0);

            RichTextBox box = new RichTextBox();
            box.Dock = DockStyle.Fill;
            box.ReadOnly = true;
            box.BackColor = Color.White;
            box.BorderStyle = BorderStyle.FixedSingle;
            box.Font = UiKit.LogFont;
            box.HideSelection = false;
            box.DetectUrls = false;
            box.WordWrap = true;

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

        // ---------------- 事件 ----------------

        private void BrowseExcel()
        {
            using (OpenFileDialog d = new OpenFileDialog())
            {
                d.Title = "选择 Excel 文件";
                d.Filter = "Excel 文件 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|所有文件 (*.*)|*.*";
                if (d.ShowDialog(this) == DialogResult.OK)
                {
                    _txtExcel.Text = d.FileName;
                    _table = null;
                    _cbPrefix.Items.Clear();
                    _cbOld.Items.Clear();
                }
            }
        }

        private void BrowseFolder()
        {
            using (FolderBrowserDialog d = new FolderBrowserDialog())
            {
                d.Description = "选择文件所在文件夹";
                if (!string.IsNullOrEmpty(_txtFolder.Text) && Directory.Exists(_txtFolder.Text))
                    d.SelectedPath = _txtFolder.Text;
                if (d.ShowDialog(this) == DialogResult.OK) _txtFolder.Text = d.SelectedPath;
            }
        }

        private void BrowseOutput()
        {
            using (FolderBrowserDialog d = new FolderBrowserDialog())
            {
                d.Description = "选择输出文件夹（改名后的文件复制到这里）";
                if (d.ShowDialog(this) == DialogResult.OK) _txtOutput.Text = d.SelectedPath;
            }
        }

        private void ToggleOutput()
        {
            bool copy = _rbOpCopy.Checked;
            if (_txtOutput != null) _txtOutput.Enabled = copy;
            if (_btnOutput != null) _btnOutput.Enabled = copy;
        }

        private void LoadExcel()
        {
            if (_busy) return;
            string path = _txtExcel.Text.Trim();
            if (!File.Exists(path))
            {
                MessageBox.Show(this, "请先选择 Excel 文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _ui.Log("正在读取 Excel：" + Path.GetFileName(path), "info");
            SetBusy(true);
            Task.Run(delegate { return ExcelTable.Load(path); })
                .ContinueWith(delegate(Task<ExcelTable> t)
                {
                    SetBusy(false);
                    if (t.IsFaulted)
                    {
                        Exception ex = t.Exception == null ? null : t.Exception.GetBaseException();
                        string msg = ex == null ? "未知错误" : ex.Message;
                        _ui.Log("读取 Excel 异常：" + msg, "miss");
                        MessageBox.Show(this, msg, "读取 Excel 失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    ExcelTable table = t.Result;
                    _table = table;
                    _cbPrefix.Items.Clear();
                    _cbOld.Items.Clear();
                    foreach (string h in table.Headers)
                    {
                        _cbPrefix.Items.Add(h);
                        _cbOld.Items.Add(h);
                    }
                    if (_cbPrefix.Items.Count > 0) _cbPrefix.SelectedIndex = 0;
                    if (_cbOld.Items.Count > 1) _cbOld.SelectedIndex = 1;
                    else if (_cbOld.Items.Count > 0) _cbOld.SelectedIndex = 0;
                    _txtEndRow.Text = (table.Rows.Count + table.HeaderRow).ToString();
                    _ui.Log(string.Format("Excel 读取成功：工作表「{0}」，表头在第 {1} 行，共 {2} 行数据（Excel 第 {3}~{4} 行）",
                        table.SheetName, table.HeaderRow, table.Rows.Count,
                        table.HeaderRow + 1, table.Rows.Count + table.HeaderRow), "ok");
                    _ui.Log("可用列：" + string.Join("  |  ", table.Headers.ToArray()), "info");
                }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private bool CollectOptions(out ExcelRenameOptions opt)
        {
            opt = null;
            if (_table == null)
            {
                MessageBox.Show(this, "请先点击【加载 Excel 列】读取表格", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            string folder = _txtFolder.Text.Trim();
            if (!Directory.Exists(folder))
            {
                MessageBox.Show(this, "请选择文件所在文件夹", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (_cbPrefix.SelectedItem == null || _cbOld.SelectedItem == null)
            {
                MessageBox.Show(this, "请选择前缀列和旧名字列", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            ExcelRenameOptions o = new ExcelRenameOptions();
            o.ExcelPath = _txtExcel.Text.Trim();
            o.Folder = folder;
            o.ColPrefix = _cbPrefix.SelectedItem.ToString();
            o.ColOldId = _cbOld.SelectedItem.ToString();
            o.Separator = _txtSeparator.Text;
            o.Extension = _txtExtension.Text.Trim();
            o.PadDigits = (int)_numPad.Value;
            o.AutoClean = _chkClean.Checked;
            o.DupMode = _rbDupAuto.Checked ? "auto" : "skip";
            o.OpMode = _rbOpCopy.Checked ? "copy" : "rename";
            o.OutputFolder = _txtOutput.Text.Trim();

            if (o.Extension.TrimStart('.').Length == 0)
            {
                MessageBox.Show(this, "请填写扩展名（不带点），如 pdf / docx", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            int start = 2, end = 0;
            string s = _txtStartRow.Text.Trim();
            string e = _txtEndRow.Text.Trim();
            if (s.Length > 0 && !int.TryParse(s, out start))
            {
                MessageBox.Show(this, "起始行必须是数字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (e.Length > 0 && !int.TryParse(e, out end))
            {
                MessageBox.Show(this, "结束行必须是数字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (start < 2) start = 2;
            if (end > 0 && end < start)
            {
                MessageBox.Show(this, "行范围无效：结束行不能小于起始行", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            o.StartRow = start;
            o.EndRow = end;

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
            ExcelRenameOptions opt;
            if (!CollectOptions(out opt)) return;

            ExcelTable table = _table;
            _ui.Log("");
            _ui.Log("======= 预览模式，不会修改文件 =======", "info");
            _ui.ResetProgress();
            SetBusy(true);
            Task.Run(delegate
            {
                return RenameEngine.PreviewExcel(table, opt,
                    delegate(LogItem li) { _ui.Log(li); },
                    delegate(int cur, int max) { _ui.Progress(cur, max); });
            }).ContinueWith(delegate(Task<PreviewResult> t)
            {
                SetBusy(false);
                _ui.ResetProgress();
                if (t.IsFaulted)
                {
                    Exception ex = t.Exception == null ? null : t.Exception.GetBaseException();
                    string msg = ex == null ? "未知错误" : ex.Message;
                    _ui.Log("预览失败：" + msg, "miss");
                    MessageBox.Show(this, msg, "预览失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                _lastPreview = t.Result;
                _lastOptions = opt;
                _ui.Log("确认全部正确后，点击【执行正式改名】", "info");
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void DoRename()
        {
            if (_busy) return;
            if (_lastPreview == null || _lastPreview.Items.Count == 0)
            {
                MessageBox.Show(this, "请先执行【预览改名】", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            ExcelRenameOptions opt;
            if (!CollectOptions(out opt)) return;

            if (_lastOptions == null || !SameOptions(opt, _lastOptions))
            {
                MessageBox.Show(this, "规则或路径已修改，请重新点击【预览改名】再执行。", "提示",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool isCopy = opt.OpMode == "copy";
            PreviewResult pv = _lastPreview;
            string warn = "";
            if (pv.MissingCount > 0 || pv.DupCount > 0)
            {
                warn = string.Format("\n\n注意：预览中发现 {0} 个文件缺失，{1} 个重名跳过。\n这些文件不会被处理，是否仍要继续？",
                    pv.MissingCount, pv.DupCount);
            }
            string msg = isCopy
                ? "确认执行复制改名！\n原文件保留不动，改名后的副本输出到：\n" + opt.OutputFolder + warn + "\n\n确认继续？"
                : "确认执行改名！\n将直接修改原文件夹中的文件名！\n强烈建议先备份文件！" + warn + "\n\n确认继续？";
            if (MessageBox.Show(this, msg, "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

            _ui.Log("");
            _ui.Log("======= " + (isCopy ? "复制改名" : "正式改名") + "开始 =======", "info");
            _ui.ResetProgress();
            SetBusy(true);
            Task.Run(delegate
            {
                return RenameEngine.ExecuteExcel(pv.Items, opt,
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
                if (r.RecordFile != null)
                    _ui.Log("如需撤销，点击【撤销改名】并选择该记录文件：" + Path.GetFileName(r.RecordFile), "info");
                MessageBox.Show(this, string.Format("执行完成：成功 {0} 个，失败 {1} 个。", r.Success, r.Failed),
                                "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private static bool SameOptions(ExcelRenameOptions a, ExcelRenameOptions b)
        {
            return a.Folder == b.Folder && a.ColPrefix == b.ColPrefix && a.ColOldId == b.ColOldId
                && a.Separator == b.Separator && a.Extension == b.Extension && a.PadDigits == b.PadDigits
                && a.AutoClean == b.AutoClean && a.DupMode == b.DupMode && a.OpMode == b.OpMode
                && a.OutputFolder == b.OutputFolder && a.StartRow == b.StartRow && a.EndRow == b.EndRow;
        }

        private void DoUndo()
        {
            if (_busy) return;
            string folder = _txtFolder.Text.Trim();
            if (!Directory.Exists(folder))
            {
                MessageBox.Show(this, "请先选择文件所在文件夹（撤销时需要）", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string fp;
            using (OpenFileDialog d = new OpenFileDialog())
            {
                d.Title = "选择改名记录文件";
                d.Filter = "CSV 记录文件 (*.csv)|*.csv|所有文件 (*.*)|*.*";
                d.InitialDirectory = folder;
                if (d.ShowDialog(this) != DialogResult.OK) return;
                fp = d.FileName;
            }
            string recordDir;
            bool isCopy;
            try
            {
                recordDir = Path.GetDirectoryName(Path.GetFullPath(fp));
                isCopy = RenameEngine.IsCopyModeRecord(fp, folder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string msg = isCopy
                ? "检测到这是复制模式生成的记录。\n将从输出文件夹删除以下副本文件：\n" + recordDir + "\n\n原文件夹中的文件不受影响。\n\n确认继续？"
                : "将根据记录文件把文件名改回原名。\n记录文件：\n" + fp + "\n\n确认继续？";
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

        private void ExportLog()
        {
            string content = _ui.AllText().Trim();
            if (content.Length == 0)
            {
                MessageBox.Show(this, "日志为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            using (SaveFileDialog d = new SaveFileDialog())
            {
                d.Title = "导出日志";
                d.Filter = "文本文件 (*.txt)|*.txt";
                d.FileName = "rename_log.txt";
                if (d.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    File.WriteAllText(d.FileName, content, new UTF8Encoding(true));
                    MessageBox.Show(this, "日志已导出到：\n" + d.FileName, "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "导出失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void SetBusy(bool busy)
        {
            _busy = busy;
            UseWaitCursor = busy;
            _btnLoad.Enabled = !busy;
            _btnPreview.Enabled = !busy;
            _btnRename.Enabled = !busy;
            _btnUndo.Enabled = !busy;
            _btnBatch.Enabled = !busy;
            _btnExport.Enabled = !busy;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _timer.Stop();
            _timer.Dispose();
            base.OnFormClosing(e);
        }
    }
}