using ExcelBatchRenamer.Models;
using ExcelBatchRenamer.Services;
using ExcelBatchRenamer.Utils;

namespace ExcelBatchRenamer;

public class MainForm : Form
{
    private readonly ExcelReader _excelReader = new();
    private readonly FileRenamer _renamer = new();
    private readonly UndoManager _undoManager = new();

    private List<string> _headers = new();
    private List<string[]> _rows = new();
    private List<RenameResult> _lastResults = new();

    // 控件
    private TextBox _txtExcel = null!;
    private TextBox _txtFolder = null!;
    private ComboBox _cbPrefix = null!;
    private ComboBox _cbOld = null!;
    private TextBox _txtStartRow = null!;
    private TextBox _txtEndRow = null!;
    private TextBox _txtSeparator = null!;
    private TextBox _txtExtension = null!;
    private NumericUpDown _numPad = null!;
    private RadioButton _rbSkip = null!;
    private RadioButton _rbAuto = null!;
    private CheckBox _chkClean = null!;
    private RadioButton _rbRename = null!;
    private RadioButton _rbCopy = null!;
    private TextBox _txtOutput = null!;
    private Button _btnOutput = null!;
    private ProgressBar _progress = null!;
    private RichTextBox _log = null!;

    public MainForm()
    {
        Text = "Excel批量文件重命名工具 Pro v2.0";
        Size = new Size(880, 720);
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(800, 600);
        Font = new Font("Microsoft YaHei UI", 9F);
        BuildUi();
    }

    private void BuildUi()
    {
        var panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        var inner = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 3,
            Padding = new Padding(10),
        };
        inner.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        inner.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        inner.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));

        int row = 0;

        // 1. Excel文件
        inner.Controls.Add(new Label { Text = "1. 选择Excel文件(.xlsx)", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft }, 0, row);
        _txtExcel = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, ReadOnly = true };
        inner.Controls.Add(_txtExcel, 1, row);
        var btnExcel = new Button { Text = "浏览", Anchor = AnchorStyles.Left };
        btnExcel.Click += (s, e) => SelectExcel();
        inner.Controls.Add(btnExcel, 2, row);
        row++;

        // 2. 文件夹
        inner.Controls.Add(new Label { Text = "2. 选择文件所在文件夹", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft }, 0, row);
        _txtFolder = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, ReadOnly = true };
        inner.Controls.Add(_txtFolder, 1, row);
        var btnFolder = new Button { Text = "浏览", Anchor = AnchorStyles.Left };
        btnFolder.Click += (s, e) => SelectFolder();
        inner.Controls.Add(btnFolder, 2, row);
        row++;

        // 3. 前缀列
        inner.Controls.Add(new Label { Text = "3. 前缀列", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft }, 0, row);
        _cbPrefix = new ComboBox { Anchor = AnchorStyles.Left, Width = 250, DropDownStyle = ComboBoxStyle.DropDownList };
        inner.Controls.Add(_cbPrefix, 1, row);
        row++;

        // 4. 旧名字列
        inner.Controls.Add(new Label { Text = "4. 旧名字列", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft }, 0, row);
        _cbOld = new ComboBox { Anchor = AnchorStyles.Left, Width = 250, DropDownStyle = ComboBoxStyle.DropDownList };
        inner.Controls.Add(_cbOld, 1, row);
        row++;

        // 5. 行范围
        inner.Controls.Add(new Label { Text = "5. 处理行范围", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft }, 0, row);
        var rangePanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, WrapContents = false, AutoSize = true };
        rangePanel.Controls.Add(new Label { Text = "从第", TextAlign = ContentAlignment.MiddleLeft });
        _txtStartRow = new TextBox { Width = 50, Text = "2" };
        rangePanel.Controls.Add(_txtStartRow);
        rangePanel.Controls.Add(new Label { Text = "行 到第", TextAlign = ContentAlignment.MiddleLeft });
        _txtEndRow = new TextBox { Width = 50 };
        rangePanel.Controls.Add(_txtEndRow);
        rangePanel.Controls.Add(new Label { Text = "行（留空=全部，行号含表头行）", TextAlign = ContentAlignment.MiddleLeft });
        inner.Controls.Add(rangePanel, 1, row);
        row++;

        // 6. 命名规则分组
        var ruleGroup = new GroupBox { Text = "命名规则设置", Anchor = AnchorStyles.Left | AnchorStyles.Right, Height = 130 };
        var ruleTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 6, RowCount = 3, Padding = new Padding(8) };
        for (int i = 0; i < 6; i++) ruleTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        ruleTable.Controls.Add(new Label { Text = "连接符:", TextAlign = ContentAlignment.MiddleRight }, 0, 0);
        _txtSeparator = new TextBox { Width = 60, Text = "_" };
        ruleTable.Controls.Add(_txtSeparator, 1, 0);
        ruleTable.Controls.Add(new Label { Text = "(前缀和旧名之间，支持中文)", TextAlign = ContentAlignment.MiddleLeft }, 2, 0);
        ruleTable.Controls.Add(new Label { Text = "扩展名:", TextAlign = ContentAlignment.MiddleRight }, 3, 0);
        _txtExtension = new TextBox { Width = 60, Text = "pdf" };
        ruleTable.Controls.Add(_txtExtension, 4, 0);
        ruleTable.Controls.Add(new Label { Text = "(不带点)", TextAlign = ContentAlignment.MiddleLeft }, 5, 0);

        ruleTable.Controls.Add(new Label { Text = "前缀补零:", TextAlign = ContentAlignment.MiddleRight }, 0, 1);
        _numPad = new NumericUpDown { Width = 60, Minimum = 0, Maximum = 10, Value = 0 };
        ruleTable.Controls.Add(_numPad, 1, 1);
        ruleTable.Controls.Add(new Label { Text = "(0=不补零，如3->001填3)", TextAlign = ContentAlignment.MiddleLeft }, 2, 1);
        ruleTable.Controls.Add(new Label { Text = "重名处理:", TextAlign = ContentAlignment.MiddleRight }, 3, 1);
        _rbSkip = new RadioButton { Text = "跳过", Checked = true };
        ruleTable.Controls.Add(_rbSkip, 4, 1);
        _rbAuto = new RadioButton { Text = "自动加序号" };
        ruleTable.Controls.Add(_rbAuto, 5, 1);

        _chkClean = new CheckBox { Text = "自动清洗文件名（去空格、替换非法字符）", Checked = true };
        ruleTable.Controls.Add(_chkClean, 0, 2);
        ruleTable.SetColumnSpan(_chkClean, 6);

        ruleGroup.Controls.Add(ruleTable);
        inner.Controls.Add(ruleGroup, 0, row);
        inner.SetColumnSpan(ruleGroup, 3);
        row++;

        // 7. 操作模式
        var modeGroup = new GroupBox { Text = "操作模式", Anchor = AnchorStyles.Left | AnchorStyles.Right, Height = 80 };
        var modeTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 2, Padding = new Padding(8) };
        _rbRename = new RadioButton { Text = "直接改名（原文件被改）", Checked = true };
        _rbRename.CheckedChanged += (s, e) => ToggleOutput();
        modeTable.Controls.Add(_rbRename, 0, 0);
        _rbCopy = new RadioButton { Text = "复制到新文件夹（原文件保留，推荐）" };
        _rbCopy.CheckedChanged += (s, e) => ToggleOutput();
        modeTable.Controls.Add(_rbCopy, 1, 0);
        modeTable.SetColumnSpan(_rbCopy, 3);

        modeTable.Controls.Add(new Label { Text = "输出文件夹:", TextAlign = ContentAlignment.MiddleRight }, 0, 1);
        _txtOutput = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, ReadOnly = true, Enabled = false };
        modeTable.Controls.Add(_txtOutput, 1, 1);
        modeTable.SetColumnSpan(_txtOutput, 2);
        _btnOutput = new Button { Text = "浏览", Enabled = false };
        _btnOutput.Click += (s, e) => SelectOutputFolder();
        modeTable.Controls.Add(_btnOutput, 3, 1);

        modeGroup.Controls.Add(modeTable);
        inner.Controls.Add(modeGroup, 0, row);
        inner.SetColumnSpan(modeGroup, 3);
        row++;

        // 8. 按钮行
        var btnPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, WrapContents = true, AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Right };
        var btnLoad = new Button { Text = "加载Excel列", Width = 100 };
        btnLoad.Click += (s, e) => LoadExcelColumns();
        btnPanel.Controls.Add(btnLoad);

        var btnPreview = new Button { Text = "预览改名(不改文件)", Width = 130 };
        btnPreview.Click += async (s, e) => await DoPreview();
        btnPanel.Controls.Add(btnPreview);

        var btnExecute = new Button { Text = "执行正式改名", Width = 110, BackColor = Color.FromArgb(220, 50, 50), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        btnExecute.FlatAppearance.BorderSize = 0;
        btnExecute.Click += async (s, e) => await DoExecute();
        btnPanel.Controls.Add(btnExecute);

        var btnUndo = new Button { Text = "撤销改名", Width = 90 };
        btnUndo.Click += (s, e) => DoUndo();
        btnPanel.Controls.Add(btnUndo);

        var btnBatch = new Button { Text = "普通批量改名", Width = 110 };
        btnBatch.Click += (s, e) => OpenBatchWindow();
        btnPanel.Controls.Add(btnBatch);

        var btnExport = new Button { Text = "导出日志", Width = 90 };
        btnExport.Click += (s, e) => ExportLog();
        btnPanel.Controls.Add(btnExport);

        var btnClear = new Button { Text = "清空日志", Width = 90 };
        btnClear.Click += (s, e) => _log.Clear();
        btnPanel.Controls.Add(btnClear);

        inner.Controls.Add(btnPanel, 0, row);
        inner.SetColumnSpan(btnPanel, 3);
        row++;

        // 9. 进度条
        _progress = new ProgressBar { Anchor = AnchorStyles.Left | AnchorStyles.Right, Height = 20, Minimum = 0, Maximum = 100 };
        inner.Controls.Add(_progress, 0, row);
        inner.SetColumnSpan(_progress, 3);
        row++;

        // 10. 日志
        inner.Controls.Add(new Label { Text = "运行日志：", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft }, 0, row);
        row++;
        _log = new RichTextBox
        {
            ReadOnly = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            Height = 250,
            Font = new Font("Consolas", 9F),
            BackColor = Color.FromArgb(30, 30, 30),
            ForeColor = Color.LightGray
        };
        inner.Controls.Add(_log, 0, row);
        inner.SetColumnSpan(_log, 3);

        panel.Controls.Add(inner);
        Controls.Add(panel);
    }

    private void Log(string msg, Color? color = null)
    {
        if (InvokeRequired)
        {
            Invoke(new Action(() => Log(msg, color)));
            return;
        }
        _log.SelectionStart = _log.TextLength;
        _log.SelectionLength = 0;
        if (color.HasValue)
            _log.SelectionColor = color.Value;
        else
            _log.SelectionColor = Color.LightGray;
        _log.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}\r\n");
        _log.ScrollToCaret();
    }

    private void SelectExcel()
    {
        using var dlg = new OpenFileDialog { Filter = "Excel文件|*.xlsx", Title = "选择Excel映射表" };
        if (dlg.ShowDialog() == DialogResult.OK)
        {
            _txtExcel.Text = dlg.FileName;
            Log($"已选择Excel: {Path.GetFileName(dlg.FileName)}");
        }
    }

    private void SelectFolder()
    {
        using var dlg = new FolderBrowserDialog { Description = "选择需要改名的文件所在文件夹" };
        if (dlg.ShowDialog() == DialogResult.OK)
        {
            _txtFolder.Text = dlg.SelectedPath;
            Log($"已选择文件夹: {dlg.SelectedPath}");
        }
    }

    private void SelectOutputFolder()
    {
        using var dlg = new FolderBrowserDialog { Description = "选择输出文件夹" };
        if (dlg.ShowDialog() == DialogResult.OK)
        {
            _txtOutput.Text = dlg.SelectedPath;
            Log($"已选择输出文件夹: {dlg.SelectedPath}");
        }
    }

    private void ToggleOutput()
    {
        bool copy = _rbCopy.Checked;
        _txtOutput.Enabled = copy;
        _btnOutput.Enabled = copy;
    }

    private void LoadExcelColumns()
    {
        if (string.IsNullOrEmpty(_txtExcel.Text))
        {
            MessageBox.Show("请先选择Excel文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        try
        {
            (_headers, _rows) = _excelReader.Read(_txtExcel.Text);
            _cbPrefix.Items.Clear();
            _cbOld.Items.Clear();
            foreach (var h in _headers)
            {
                _cbPrefix.Items.Add(h);
                _cbOld.Items.Add(h);
            }
            if (_headers.Count > 0)
            {
                _cbPrefix.SelectedIndex = 0;
                _cbOld.SelectedIndex = Math.Min(1, _headers.Count - 1);
            }
            Log($"Excel加载成功：{_headers.Count}列，{_rows.Count}行数据");
        }
        catch (Exception ex)
        {
            Log($"Excel读取失败: {ex.Message}", Color.Red);
            MessageBox.Show($"Excel读取失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private RenameConfig? BuildConfig()
    {
        if (string.IsNullOrEmpty(_txtFolder.Text))
        {
            MessageBox.Show("请选择文件所在文件夹", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }
        if (_cbPrefix.SelectedIndex < 0 || _cbOld.SelectedIndex < 0)
        {
            MessageBox.Show("请先加载Excel并选择前缀列和旧名字列", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }
        if (_rbCopy.Checked && string.IsNullOrEmpty(_txtOutput.Text))
        {
            MessageBox.Show("复制模式下请选择输出文件夹", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }

        int startRow = int.TryParse(_txtStartRow.Text, out int s) ? s : 2;
        int endRow = int.TryParse(_txtEndRow.Text, out int e) ? e : 0;

        var items = _excelReader.ExtractItems(_headers, _rows, _cbPrefix.SelectedIndex, _cbOld.SelectedIndex, startRow, endRow);

        return new RenameConfig
        {
            SourceFolder = _txtFolder.Text,
            OutputFolder = _txtOutput.Text,
            Items = items,
            Separator = _txtSeparator.Text,
            Extension = _txtExtension.Text,
            PadDigits = (int)_numPad.Value,
            DupMode = _rbAuto.Checked ? DuplicateMode.AutoNumber : DuplicateMode.Skip,
            Mode = _rbCopy.Checked ? OperationMode.Copy : OperationMode.Rename,
            AutoClean = _chkClean.Checked
        };
    }

    private async Task DoPreview()
    {
        var config = BuildConfig();
        if (config == null) return;

        Log($"开始预览：共 {config.Items.Count} 条映射");
        _progress.Value = 0;

        var results = await Task.Run(() => _renamer.Preview(config));
        _lastResults = results;

        int ok = 0, miss = 0, dup = 0;
        foreach (var r in results)
        {
            if (r.Success)
            {
                ok++;
                Log($"  [OK] 行{r.RowNumber}: {r.OldName} -> {r.NewName}", Color.LightGreen);
            }
            else if (r.Message == "文件缺失")
            {
                miss++;
                Log($"  [缺失] 行{r.RowNumber}: {r.OldName}", Color.OrangeRed);
            }
            else
            {
                dup++;
                Log($"  [跳过] 行{r.RowNumber}: {r.OldName} ({r.Message})", Color.Orange);
            }
        }
        _progress.Value = 100;
        Log($"预览完成：成功 {ok}，缺失 {miss}，跳过 {dup}");
        if (miss > 0)
            Log($"⚠ 有 {miss} 个文件在磁盘上找不到，请核对Excel中的文件名", Color.Yellow);
    }

    private async Task DoExecute()
    {
        var config = BuildConfig();
        if (config == null) return;

        var confirm = MessageBox.Show(
            $"确认执行正式改名？\n\n模式: {(config.Mode == OperationMode.Copy ? "复制到新文件夹" : "直接改名")}\n共 {config.Items.Count} 条映射\n\n建议先使用预览功能确认无误。",
            "确认执行", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (confirm != DialogResult.Yes) return;

        Log($"开始执行：{config.Mode}模式，共 {config.Items.Count} 条");
        _progress.Value = 0;

        var progress = new Progress<int>(v => _progress.Value = v);
        var results = await Task.Run(() => _renamer.Execute(config, progress));
        _lastResults = results;

        int ok = 0, fail = 0;
        foreach (var r in results)
        {
            if (r.Success)
            {
                ok++;
                Log($"  [OK] {r.OldName} -> {r.NewName}", Color.LightGreen);
            }
            else
            {
                fail++;
                Log($"  [失败] {r.OldName}: {r.Message}", Color.Red);
            }
        }
        _progress.Value = 100;
        Log($"执行完成：成功 {ok}，失败 {fail}");

        // 保存撤销记录
        if (ok > 0)
        {
            try
            {
                string recordFile = _undoManager.SaveRecord(config.TargetFolder, results, config.Mode);
                Log($"撤销记录已保存: {Path.GetFileName(recordFile)}", Color.Cyan);
            }
            catch (Exception ex)
            {
                Log($"保存撤销记录失败: {ex.Message}", Color.Red);
            }
        }
    }

    private void DoUndo()
    {
        using var dlg = new OpenFileDialog
        {
            Filter = "撤销记录|rename_record_*.csv",
            Title = "选择撤销记录文件",
            InitialDirectory = _txtFolder.Text
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        var confirm = MessageBox.Show("确认执行撤销？将把文件名改回原名（或删除复制的副本）。", "确认撤销", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (confirm != DialogResult.Yes) return;

        Log($"开始撤销: {Path.GetFileName(dlg.FileName)}");
        var (success, failed, message) = _undoManager.Undo(dlg.FileName);
        Log(message, success > 0 ? Color.LightGreen : Color.Orange);
    }

    private void OpenBatchWindow()
    {
        using var batch = new BatchRenameForm();
        batch.ShowDialog(this);
    }

    private void ExportLog()
    {
        if (string.IsNullOrEmpty(_log.Text))
        {
            MessageBox.Show("日志为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        using var dlg = new SaveFileDialog
        {
            Filter = "文本文件|*.txt",
            FileName = $"rename_log_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
        };
        if (dlg.ShowDialog() == DialogResult.OK)
        {
            File.WriteAllText(dlg.FileName, _log.Text, System.Text.Encoding.UTF8);
            Log($"日志已导出: {dlg.FileName}", Color.Cyan);
        }
    }
}
