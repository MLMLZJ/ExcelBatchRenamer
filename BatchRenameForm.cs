using System.Text.RegularExpressions;
using ExcelBatchRenamer.Models;
using ExcelBatchRenamer.Services;
using ExcelBatchRenamer.Utils;

namespace ExcelBatchRenamer;

public class BatchRenameForm : Form
{
    private readonly UndoManager _undoManager = new();
    private List<(string oldPath, string newName)> _previewItems = new();

    private TextBox _txtFolder = null!;
    private TextBox _txtFilter = null!;
    private TextBox _txtPrefix = null!;
    private TextBox _txtSuffix = null!;
    private TextBox _txtFind = null!;
    private TextBox _txtReplace = null!;
    private CheckBox _chkRegex = null!;
    private ComboBox _cbCase = null!;
    private RadioButton _rbRename = null!;
    private RadioButton _rbCopy = null!;
    private TextBox _txtOutput = null!;
    private Button _btnOutput = null!;
    private RadioButton _rbSkip = null!;
    private RadioButton _rbAuto = null!;
    private ProgressBar _progress = null!;
    private RichTextBox _log = null!;

    public BatchRenameForm()
    {
        Text = "普通批量改名";
        Size = new Size(720, 640);
        StartPosition = FormStartPosition.CenterParent;
        MinimumSize = new Size(640, 560);
        Font = new Font("Microsoft YaHei UI", 9F);
        BuildUi();
    }

    private void BuildUi()
    {
        var panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 3,
            Padding = new Padding(10),
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70));

        int row = 0;

        // 文件夹
        table.Controls.Add(new Label { Text = "选择文件夹", TextAlign = ContentAlignment.MiddleLeft }, 0, row);
        _txtFolder = new TextBox { ReadOnly = true, Anchor = AnchorStyles.Left | AnchorStyles.Right };
        table.Controls.Add(_txtFolder, 1, row);
        var btnFolder = new Button { Text = "浏览" };
        btnFolder.Click += (s, e) => SelectFolder();
        table.Controls.Add(btnFolder, 2, row);
        row++;

        // 文件类型过滤
        table.Controls.Add(new Label { Text = "文件类型过滤", TextAlign = ContentAlignment.MiddleLeft }, 0, row);
        _txtFilter = new TextBox { Text = "*.*", Anchor = AnchorStyles.Left | AnchorStyles.Right };
        table.Controls.Add(_txtFilter, 1, row);
        table.Controls.Add(new Label { Text = "多个用空格", TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.Gray }, 2, row);
        row++;

        // 规则分组
        var ruleGroup = new GroupBox { Text = "改名规则（顺序：查找替换→大小写→前缀→后缀）", Anchor = AnchorStyles.Left | AnchorStyles.Right, Height = 180 };
        var ruleTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 4, Padding = new Padding(8) };

        ruleTable.Controls.Add(new Label { Text = "添加前缀:", TextAlign = ContentAlignment.MiddleRight }, 0, 0);
        _txtPrefix = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
        ruleTable.Controls.Add(_txtPrefix, 1, 0);
        ruleTable.Controls.Add(new Label { Text = "添加后缀:", TextAlign = ContentAlignment.MiddleRight }, 2, 0);
        _txtSuffix = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
        ruleTable.Controls.Add(_txtSuffix, 3, 0);

        ruleTable.Controls.Add(new Label { Text = "查找:", TextAlign = ContentAlignment.MiddleRight }, 0, 1);
        _txtFind = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
        ruleTable.Controls.Add(_txtFind, 1, 1);
        ruleTable.Controls.Add(new Label { Text = "替换为:", TextAlign = ContentAlignment.MiddleRight }, 2, 1);
        _txtReplace = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
        ruleTable.Controls.Add(_txtReplace, 3, 1);

        _chkRegex = new CheckBox { Text = "使用正则表达式" };
        ruleTable.Controls.Add(_chkRegex, 1, 2);

        ruleTable.Controls.Add(new Label { Text = "大小写转换:", TextAlign = ContentAlignment.MiddleRight }, 2, 2);
        _cbCase = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120 };
        _cbCase.Items.AddRange(new object[] { "不变", "全部大写", "全部小写", "首字母大写" });
        _cbCase.SelectedIndex = 0;
        ruleTable.Controls.Add(_cbCase, 3, 2);

        ruleGroup.Controls.Add(ruleTable);
        table.Controls.Add(ruleGroup, 0, row);
        table.SetColumnSpan(ruleGroup, 3);
        row++;

        // 操作模式
        var modeGroup = new GroupBox { Text = "操作模式", Anchor = AnchorStyles.Left | AnchorStyles.Right, Height = 80 };
        var modeTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 2, Padding = new Padding(8) };
        _rbRename = new RadioButton { Text = "直接改名", Checked = true };
        _rbRename.CheckedChanged += (s, e) => ToggleOutput();
        modeTable.Controls.Add(_rbRename, 0, 0);
        _rbCopy = new RadioButton { Text = "复制到新文件夹（推荐）" };
        _rbCopy.CheckedChanged += (s, e) => ToggleOutput();
        modeTable.Controls.Add(_rbCopy, 1, 0);
        modeTable.SetColumnSpan(_rbCopy, 3);

        modeTable.Controls.Add(new Label { Text = "输出文件夹:", TextAlign = ContentAlignment.MiddleRight }, 0, 1);
        _txtOutput = new TextBox { ReadOnly = true, Enabled = false, Anchor = AnchorStyles.Left | AnchorStyles.Right };
        modeTable.Controls.Add(_txtOutput, 1, 1);
        modeTable.SetColumnSpan(_txtOutput, 2);
        _btnOutput = new Button { Text = "浏览", Enabled = false };
        _btnOutput.Click += (s, e) => SelectOutput();
        modeTable.Controls.Add(_btnOutput, 3, 1);
        modeGroup.Controls.Add(modeTable);
        table.Controls.Add(modeGroup, 0, row);
        table.SetColumnSpan(modeGroup, 3);
        row++;

        // 重名处理
        var dupGroup = new GroupBox { Text = "重名处理", Anchor = AnchorStyles.Left, Width = 300, Height = 50 };
        var dupPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };
        _rbSkip = new RadioButton { Text = "跳过", Checked = true };
        _rbAuto = new RadioButton { Text = "自动加序号(1)(2)" };
        dupPanel.Controls.Add(_rbSkip);
        dupPanel.Controls.Add(_rbAuto);
        dupGroup.Controls.Add(dupPanel);
        table.Controls.Add(dupGroup, 0, row);
        row++;

        // 按钮
        var btnPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Right };
        var btnPreview = new Button { Text = "预览", Width = 80 };
        btnPreview.Click += (s, e) => DoPreview();
        btnPanel.Controls.Add(btnPreview);
        var btnExecute = new Button { Text = "执行改名", Width = 90, BackColor = Color.FromArgb(220, 50, 50), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        btnExecute.FlatAppearance.BorderSize = 0;
        btnExecute.Click += (s, e) => DoExecute();
        btnPanel.Controls.Add(btnExecute);
        var btnUndo = new Button { Text = "撤销", Width = 80 };
        btnUndo.Click += (s, e) => DoUndo();
        btnPanel.Controls.Add(btnUndo);
        var btnClose = new Button { Text = "关闭", Width = 80 };
        btnClose.Click += (s, e) => Close();
        btnPanel.Controls.Add(btnClose);
        table.Controls.Add(btnPanel, 0, row);
        table.SetColumnSpan(btnPanel, 3);
        row++;

        // 进度条
        _progress = new ProgressBar { Anchor = AnchorStyles.Left | AnchorStyles.Right, Height = 18 };
        table.Controls.Add(_progress, 0, row);
        table.SetColumnSpan(_progress, 3);
        row++;

        // 日志
        table.Controls.Add(new Label { Text = "日志：", TextAlign = ContentAlignment.MiddleLeft }, 0, row);
        row++;
        _log = new RichTextBox
        {
            ReadOnly = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            Height = 200,
            Font = new Font("Consolas", 9F),
            BackColor = Color.FromArgb(30, 30, 30),
            ForeColor = Color.LightGray
        };
        table.Controls.Add(_log, 0, row);
        table.SetColumnSpan(_log, 3);

        panel.Controls.Add(table);
        Controls.Add(panel);
    }

    private void Log(string msg, Color? color = null)
    {
        if (InvokeRequired) { Invoke(new Action(() => Log(msg, color))); return; }
        _log.SelectionStart = _log.TextLength;
        _log.SelectionColor = color ?? Color.LightGray;
        _log.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}\r\n");
        _log.ScrollToCaret();
    }

    private void SelectFolder()
    {
        using var dlg = new FolderBrowserDialog { Description = "选择需要批量改名的文件夹" };
        if (dlg.ShowDialog() == DialogResult.OK)
        {
            _txtFolder.Text = dlg.SelectedPath;
            Log($"已选择: {dlg.SelectedPath}");
        }
    }

    private void SelectOutput()
    {
        using var dlg = new FolderBrowserDialog { Description = "选择输出文件夹" };
        if (dlg.ShowDialog() == DialogResult.OK)
        {
            _txtOutput.Text = dlg.SelectedPath;
            Log($"输出文件夹: {dlg.SelectedPath}");
        }
    }

    private void ToggleOutput()
    {
        bool copy = _rbCopy.Checked;
        _txtOutput.Enabled = copy;
        _btnOutput.Enabled = copy;
    }

    private List<string> GetFiles()
    {
        if (string.IsNullOrEmpty(_txtFolder.Text) || !Directory.Exists(_txtFolder.Text))
            return new List<string>();

        var filters = _txtFilter.Text.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (filters.Length == 0) filters = new[] { "*.*" };

        var allFiles = Directory.GetFiles(_txtFolder.Text);
        var result = new List<string>();
        foreach (var file in allFiles)
        {
            string name = Path.GetFileName(file);
            foreach (var filter in filters)
            {
                if (MatchesPattern(name, filter))
                {
                    result.Add(file);
                    break;
                }
            }
        }
        return result;
    }

    private static bool MatchesPattern(string fileName, string pattern)
    {
        if (pattern == "*.*" || pattern == "*") return true;
        string regex = "^" + Regex.Escape(pattern).Replace("\\*", ".*").Replace("\\?", ".") + "$";
        return Regex.IsMatch(fileName, regex, RegexOptions.IgnoreCase);
    }

    private string ApplyRules(string fileName)
    {
        string name = Path.GetFileNameWithoutExtension(fileName);
        string ext = Path.GetExtension(fileName);

        // 1. 查找替换
        if (!string.IsNullOrEmpty(_txtFind.Text))
        {
            if (_chkRegex.Checked)
            {
                try { name = Regex.Replace(name, _txtFind.Text, _txtReplace.Text); }
                catch { /* 正则错误时跳过 */ }
            }
            else
            {
                name = name.Replace(_txtFind.Text, _txtReplace.Text);
            }
        }

        // 2. 大小写转换
        switch (_cbCase.SelectedIndex)
        {
            case 1: name = name.ToUpper(); break;
            case 2: name = name.ToLower(); break;
            case 3:
                if (name.Length > 0)
                    name = char.ToUpper(name[0]) + name.Substring(1).ToLower();
                break;
        }

        // 3. 添加前缀
        name = _txtPrefix.Text + name;

        // 4. 添加后缀
        name = name + _txtSuffix.Text;

        // 清洗
        name = FileNameCleaner.Clean(name);

        return name + ext;
    }

    private void DoPreview()
    {
        var files = GetFiles();
        if (files.Count == 0)
        {
            MessageBox.Show("没有找到匹配的文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        Log($"开始预览：共 {files.Count} 个文件");
        _previewItems.Clear();
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int ok = 0, skip = 0;

        foreach (var file in files)
        {
            string newName = ApplyRules(Path.GetFileName(file));
            if (usedNames.Contains(newName))
            {
                if (_rbSkip.Checked)
                {
                    Log($"  [跳过] {Path.GetFileName(file)} (重名)", Color.Orange);
                    skip++;
                    continue;
                }
                newName = AddNumberSuffix(newName, usedNames);
            }
            usedNames.Add(newName);
            _previewItems.Add((file, newName));
            Log($"  [OK] {Path.GetFileName(file)} -> {newName}", Color.LightGreen);
            ok++;
        }
        Log($"预览完成：{ok} 个将改名，{skip} 个跳过");
    }

    private void DoExecute()
    {
        if (_previewItems.Count == 0)
        {
            MessageBox.Show("请先点击预览", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        bool copy = _rbCopy.Checked;
        if (copy && string.IsNullOrEmpty(_txtOutput.Text))
        {
            MessageBox.Show("复制模式请选择输出文件夹", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var confirm = MessageBox.Show($"确认执行？共 {_previewItems.Count} 个文件", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (confirm != DialogResult.Yes) return;

        string targetFolder = copy ? _txtOutput.Text : _txtFolder.Text;
        if (copy) Directory.CreateDirectory(targetFolder);

        Log($"开始执行：{(copy ? "复制" : "改名")} {_previewItems.Count} 个文件");
        _progress.Value = 0;
        int ok = 0, fail = 0;
        var results = new List<RenameResult>();

        for (int i = 0; i < _previewItems.Count; i++)
        {
            var (oldPath, newName) = _previewItems[i];
            string newPath = Path.Combine(targetFolder, newName);
            try
            {
                if (copy)
                    File.Copy(oldPath, newPath, overwrite: false);
                else
                    File.Move(oldPath, newPath);

                results.Add(new RenameResult
                {
                    OldPath = oldPath,
                    NewPath = newPath,
                    OldName = Path.GetFileName(oldPath),
                    NewName = newName,
                    Success = true,
                    Message = "OK"
                });
                ok++;
                Log($"  [OK] {Path.GetFileName(oldPath)} -> {newName}", Color.LightGreen);
            }
            catch (Exception ex)
            {
                results.Add(new RenameResult
                {
                    OldPath = oldPath,
                    OldName = Path.GetFileName(oldPath),
                    NewName = newName,
                    Success = false,
                    Message = ex.Message
                });
                fail++;
                Log($"  [失败] {Path.GetFileName(oldPath)}: {ex.Message}", Color.Red);
            }
            _progress.Value = (i + 1) * 100 / _previewItems.Count;
        }
        Log($"执行完成：成功 {ok}，失败 {fail}");

        if (ok > 0)
        {
            try
            {
                string recordFile = Path.Combine(targetFolder, $"batch_record_{DateTime.Now:yyyyMMdd_HHmmss}.csv");
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("mode,old_path,new_path");
                foreach (var r in results.Where(x => x.Success))
                    sb.AppendLine($"{(copy ? "copy" : "rename")},{EscapeCsv(r.OldPath)},{EscapeCsv(r.NewPath)}");
                File.WriteAllText(recordFile, sb.ToString(), System.Text.Encoding.UTF8);
                Log($"撤销记录: {Path.GetFileName(recordFile)}", Color.Cyan);
            }
            catch (Exception ex)
            {
                Log($"保存记录失败: {ex.Message}", Color.Red);
            }
        }
        _previewItems.Clear();
    }

    private void DoUndo()
    {
        using var dlg = new OpenFileDialog
        {
            Filter = "撤销记录|*_record_*.csv",
            Title = "选择撤销记录",
            InitialDirectory = _txtFolder.Text
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        var lines = File.ReadAllLines(dlg.FileName, System.Text.Encoding.UTF8).Skip(1).ToList();
        int ok = 0, fail = 0;
        for (int i = lines.Count - 1; i >= 0; i--)
        {
            var parts = ParseCsv(lines[i]);
            if (parts.Length < 3) continue;
            string mode = parts[0], oldPath = parts[1], newPath = parts[2];
            try
            {
                if (mode == "copy")
                {
                    if (File.Exists(newPath)) { File.Delete(newPath); ok++; }
                }
                else
                {
                    if (File.Exists(newPath)) { File.Move(newPath, oldPath); ok++; }
                }
            }
            catch { fail++; }
        }
        Log($"撤销完成：成功 {ok}，失败 {fail}");
        MessageBox.Show($"撤销完成：成功 {ok}，失败 {fail}", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private static string AddNumberSuffix(string fileName, HashSet<string> used)
    {
        string name = Path.GetFileNameWithoutExtension(fileName);
        string ext = Path.GetExtension(fileName);
        int n = 1;
        string candidate;
        do { candidate = $"{name}({n}){ext}"; n++; } while (used.Contains(candidate));
        return candidate;
    }

    private static string EscapeCsv(string v) => v.Contains(',') || v.Contains('"') ? "\"" + v.Replace("\"", "\"\"") + "\"" : v;
    private static string[] ParseCsv(string line)
    {
        var result = new List<string>();
        bool inQ = false;
        var cur = new System.Text.StringBuilder();
        foreach (char c in line)
        {
            if (inQ) { if (c == '"') inQ = false; else cur.Append(c); }
            else if (c == '"') inQ = true;
            else if (c == ',') { result.Add(cur.ToString()); cur.Clear(); }
            else cur.Append(c);
        }
        result.Add(cur.ToString());
        return result.ToArray();
    }
}
