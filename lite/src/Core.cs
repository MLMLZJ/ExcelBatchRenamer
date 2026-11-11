using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ExcelBatchRename
{
    /// <summary>一条日志</summary>
    public class LogItem
    {
        public string Text;
        public string Tag;   // null / "ok" / "miss" / "dup" / "info"
        public LogItem(string text, string tag) { Text = text; Tag = tag; }
    }

    /// <summary>一条待改名/已改名的记录</summary>
    public class RenamePair
    {
        public string OldFull;   // 源文件完整路径
        public string OldName;   // 源文件名
        public string NewName;   // 新文件名
        public RenamePair() { }
        public RenamePair(string oldFull, string oldName, string newName)
        { OldFull = oldFull; OldName = oldName; NewName = newName; }
    }

    public class PreviewResult
    {
        public List<RenamePair> Items = new List<RenamePair>();
        public List<LogItem> Log = new List<LogItem>();
        public int OkCount;
        public int MissingCount;
        public int DupCount;
        /// <summary>实时回调（界面用来逐行显示）</summary>
        public Action<LogItem> Sink;
        public void AddLog(string text, string tag = null)
        {
            LogItem item = new LogItem(text, tag);
            Log.Add(item);
            if (Sink != null) Sink(item);
        }
    }

    public class ExecuteResult
    {
        public int Success;
        public int Failed;
        public string RecordFile;
        public List<LogItem> Log = new List<LogItem>();
        /// <summary>实时回调（界面用来逐行显示）</summary>
        public Action<LogItem> Sink;
        public void AddLog(string text, string tag = null)
        {
            LogItem item = new LogItem(text, tag);
            Log.Add(item);
            if (Sink != null) Sink(item);
        }
    }

    /// <summary>读出来的 Excel 表（只读第一个工作表）</summary>
    public class ExcelTable
    {
        public string SheetName = "";
        /// <summary>表头所在的 Excel 行号（1 基）</summary>
        public int HeaderRow = 1;
        public List<string> Headers = new List<string>();
        public List<string[]> Rows = new List<string[]>();

        public int ColumnIndex(string header)
        {
            if (header == null) return -1;
            for (int i = 0; i < Headers.Count; i++)
            {
                if (Headers[i] == header) return i;
            }
            return -1;
        }

        public string Get(int rowIndex, string header)
        {
            int ci = ColumnIndex(header);
            if (ci < 0 || rowIndex < 0 || rowIndex >= Rows.Count) return "";
            string[] r = Rows[rowIndex];
            if (ci >= r.Length) return "";
            return r[ci] ?? "";
        }

        // ---------------- 载入 ----------------

        public static ExcelTable Load(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("未指定 Excel 文件");
            if (!File.Exists(filePath)) throw new FileNotFoundException("找不到 Excel 文件：" + filePath);
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext == ".xls")
                throw new NotSupportedException("不支持旧版 .xls 文件。请用 Excel 打开后“另存为” .xlsx 再试。");
            if (ext != ".xlsx" && ext != ".xlsm")
                throw new NotSupportedException("只支持 .xlsx / .xlsm 文件，当前文件：" + Path.GetFileName(filePath));

            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                List<string> shared = LoadSharedStrings(zip);

                string sheetName;
                string resolved = ResolveFirstSheet(zip, out sheetName);
                string sheetPath = resolved ?? "xl/worksheets/sheet1.xml";

                ZipArchiveEntry sheetEntry = FindEntry(zip, sheetPath);
                if (sheetEntry == null) sheetEntry = FindEntry(zip, "xl/worksheets/sheet1.xml");
                if (sheetEntry == null) throw new InvalidDataException("不是有效的 xlsx 文件（缺少工作表数据）。");

                ExcelTable table = new ExcelTable();
                table.SheetName = sheetName;

                using (Stream s = sheetEntry.Open())
                {
                    XDocument doc = XDocument.Load(s);
                    XNamespace ns = doc.Root.Name.Namespace;
                    XElement sheetData = doc.Root.Element(ns + "sheetData");
                    if (sheetData == null) throw new InvalidDataException("工作表为空或格式损坏。");

                    int expectedRow = 1;
                    bool first = true;

                    foreach (XElement rowEl in sheetData.Elements(ns + "row"))
                    {
                        XAttribute rAttr = rowEl.Attribute("r");
                        int rowNum = expectedRow;
                        if (rAttr != null)
                        {
                            int parsed;
                            if (int.TryParse(rAttr.Value, out parsed)) rowNum = parsed;
                        }
                        expectedRow = rowNum + 1;

                        SortedDictionary<int, string> cells = new SortedDictionary<int, string>();
                        int nextCol = 0;
                        foreach (XElement cEl in rowEl.Elements(ns + "c"))
                        {
                            int col = nextCol;
                            XAttribute cAttr = cEl.Attribute("r");
                            if (cAttr != null)
                            {
                                int parsed = ColFromRef(cAttr.Value);
                                if (parsed >= 0) col = parsed;
                            }
                            nextCol = col + 1;
                            cells[col] = CellText(cEl, shared, ns);
                        }

                        int maxCol = -1;
                        foreach (int k in cells.Keys) if (k > maxCol) maxCol = k;
                        string[] arr = new string[maxCol + 1];
                        for (int i = 0; i < arr.Length; i++) arr[i] = "";

                        // 补齐中间空行（表格中间没有数据的行也要占位）
                        while (table.Rows.Count < rowNum - 2)
                        {
                            table.Rows.Add(new string[0]);
                        }
                        foreach (KeyValuePair<int, string> kv in cells) arr[kv.Key] = kv.Value;

                        if (first)
                        {
                            // 与 pandas header=0 一致：Excel 第 1 行恒为表头（即使整行为空）
                            first = false;
                            table.HeaderRow = rowNum;
                            foreach (string cell in arr) table.Headers.Add(cell);
                        }
                        else
                        {
                            table.Rows.Add(arr);
                        }
                    }
                }

                // 表头为空时用占位名；重名表头加 (2)(3) 后缀，否则第二列永远选不中
                Dictionary<string, int> headerSeen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < table.Headers.Count; i++)
                {
                    string h = table.Headers[i] ?? "";
                    if (h.Length == 0) h = "列" + ColName(i);
                    int n;
                    if (headerSeen.TryGetValue(h, out n))
                    {
                        n = n + 1;
                        headerSeen[h] = n;
                        table.Headers[i] = h + "(" + n + ")";
                    }
                    else
                    {
                        headerSeen[h] = 1;
                        table.Headers[i] = h;
                    }
                }
                return table;
            }
        }

        private static ZipArchiveEntry FindEntry(ZipArchive zip, string name)
        {
            foreach (ZipArchiveEntry e in zip.Entries)
            {
                if (string.Equals(e.FullName.Replace('\\', '/'), name, StringComparison.OrdinalIgnoreCase)) return e;
            }
            return null;
        }

        private static List<string> LoadSharedStrings(ZipArchive zip)
        {
            List<string> list = new List<string>();
            ZipArchiveEntry e = FindEntry(zip, "xl/sharedStrings.xml");
            if (e == null) return list;
            using (Stream s = e.Open())
            {
                XDocument doc = XDocument.Load(s);
                XNamespace ns = doc.Root.Name.Namespace;
                foreach (XElement si in doc.Root.Elements(ns + "si"))
                {
                    list.Add(TextRuns(si, ns));
                }
            }
            return list;
        }

        /// <summary>
        /// 取 si / is 里的正文。
        /// 不能用 Descendants("t")：那会把 rPh（日文/中文注音、拼音）里的 t 也拼进来，
        /// 开了“显示拼音”的表格会读出脏字符串。
        /// </summary>
        private static string TextRuns(XElement container, XNamespace ns)
        {
            StringBuilder sb = new StringBuilder();
            foreach (XElement child in container.Elements())
            {
                if (child.Name == ns + "t") sb.Append(child.Value);
                else if (child.Name == ns + "r")
                {
                    foreach (XElement t in child.Elements(ns + "t")) sb.Append(t.Value);
                }
                // rPh / phoneticPr / rPr 一律忽略
            }
            return sb.ToString();
        }

        private static string ResolveFirstSheet(ZipArchive zip, out string sheetName)
        {
            sheetName = "";
            try
            {
                ZipArchiveEntry wb = FindEntry(zip, "xl/workbook.xml");
                ZipArchiveEntry rel = FindEntry(zip, "xl/_rels/workbook.xml.rels");
                if (wb == null) return null;

                string rid = null;
                using (Stream s = wb.Open())
                {
                    XDocument doc = XDocument.Load(s);
                    XNamespace ns = doc.Root.Name.Namespace;
                    XElement sheets = doc.Root.Element(ns + "sheets");
                    if (sheets == null) return null;
                    XElement first = sheets.Elements(ns + "sheet").FirstOrDefault();
                    if (first == null) return null;
                    sheetName = (string)first.Attribute("name") ?? "";
                    XAttribute ridAttr = first.Attributes().FirstOrDefault(a => a.Name.LocalName == "id");
                    if (ridAttr != null) rid = ridAttr.Value;
                }
                if (rid == null || rel == null) return null;

                using (Stream s = rel.Open())
                {
                    XDocument doc = XDocument.Load(s);
                    XNamespace ns = doc.Root.Name.Namespace;
                    foreach (XElement e in doc.Root.Elements(ns + "Relationship"))
                    {
                        if ((string)e.Attribute("Id") == rid)
                        {
                            string target = (string)e.Attribute("Target") ?? "";
                            target = target.Replace('\\', '/');
                            if (target.StartsWith("/")) return target.TrimStart('/');
                            return ("xl/" + target).Replace("/./", "/");
                        }
                    }
                }
            }
            catch
            {
                // 解析失败就走默认路径
            }
            return null;
        }

        private static string CellText(XElement c, List<string> shared, XNamespace ns)
        {
            string type = (string)c.Attribute("t");
            if (type == "inlineStr")
            {
                XElement isEl = c.Element(ns + "is");
                if (isEl == null) return "";
                return TextRuns(isEl, ns);
            }
            XElement vEl = c.Element(ns + "v");
            if (vEl == null) return "";
            string raw = vEl.Value;
            if (type == "s")
            {
                int idx;
                if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out idx)
                    && idx >= 0 && idx < shared.Count)
                    return shared[idx];
                return "";
            }
            if (type == "b") return raw == "1" ? "TRUE" : "FALSE";
            if (type == "str") return raw;
            return NormalizeNumber(raw);
        }

        private static string NormalizeNumber(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return raw;
            string s = raw.Trim();
            if (s.Length == 0) return s;
            if (s.IndexOf('.') >= 0 && s.IndexOf('E') < 0 && s.IndexOf('e') < 0)
            {
                string trimmed = s.TrimEnd('0');
                if (trimmed.EndsWith(".")) trimmed = trimmed.Substring(0, trimmed.Length - 1);
                if (trimmed.Length > 0) return trimmed;
            }
            return s;
        }

        /// <summary>"B12" -> 1 (0 基列号)</summary>
        public static int ColFromRef(string reference)
        {
            if (string.IsNullOrEmpty(reference)) return -1;
            int col = 0;
            bool any = false;
            foreach (char ch in reference)
            {
                char u = char.ToUpperInvariant(ch);
                if (u >= 'A' && u <= 'Z')
                {
                    col = col * 26 + (u - 'A' + 1);
                    any = true;
                }
                else break;
            }
            return any ? col - 1 : -1;
        }

        public static string ColName(int index)
        {
            string s = "";
            int n = index + 1;
            while (n > 0)
            {
                int rem = (n - 1) % 26;
                s = (char)('A' + rem) + s;
                n = (n - 1) / 26;
            }
            return s;
        }
    }

    // ---------------- 命名规则 ----------------

    public static class NameRules
    {
        private static readonly Regex Illegal = new Regex("[\\\\/:*?\"<>|]", RegexOptions.Compiled);

        public static string Clean(string name)
        {
            if (name == null) return "";
            return Illegal.Replace(name.Trim(), "_");
        }

        public static string StripExt(string name, string ext)
        {
            if (name == null) return "";
            string e = (ext ?? "").Trim().TrimStart('.').ToLowerInvariant();
            if (e.Length == 0) return name;
            if (name.ToLowerInvariant().EndsWith("." + e))
                return name.Substring(0, name.Length - (e.Length + 1));
            return name;
        }

        public static string PadPrefix(string prefix, int digits)
        {
            if (digits <= 0) return prefix;
            int num;
            if (int.TryParse(prefix, NumberStyles.Integer, CultureInfo.InvariantCulture, out num))
            {
                if (num < 0)
                {
                    // 与 Python 的 zfill 一致：符号占一位，-1 补 3 位得到 -01（总长 3），而不是 -001
                    int body = digits - 1;
                    if (body < 0) body = 0;
                    return "-" + (-(long)num).ToString(CultureInfo.InvariantCulture).PadLeft(body, '0');
                }
                return num.ToString(CultureInfo.InvariantCulture).PadLeft(digits, '0');
            }
            return prefix;
        }

        public static string BuildNewName(string prefix, string oldId, string separator, string extension,
                                          int padDigits, bool autoClean)
        {
            string ext = (extension ?? "").Trim().TrimStart('.');
            string sep = separator ?? "";
            prefix = PadPrefix((prefix ?? "").Trim(), padDigits);
            oldId = StripExt((oldId ?? "").Trim(), ext);
            if (autoClean)
            {
                prefix = Clean(prefix);
                oldId = Clean(oldId);
                sep = Clean(sep);
            }
            return prefix + sep + oldId + "." + ext;
        }

        /// <summary>把 Excel 旧名字还原成磁盘上的文件名</summary>
        public static string OldFileName(string oldId, string extension, bool autoClean)
        {
            string ext = (extension ?? "").Trim().TrimStart('.');
            string name = StripExt((oldId ?? "").Trim(), ext) + "." + ext;
            if (autoClean) name = Clean(name);
            return name;
        }

        /// <summary>
        /// 拆主名和扩展名。与 Python 的 os.path.splitext 一致：
        /// ".gitignore" 整个算主名，扩展名为空。
        /// </summary>
        public static void SplitName(string fileName, out string name, out string ext)
        {
            name = Path.GetFileNameWithoutExtension(fileName);
            ext = Path.GetExtension(fileName);
            if (name.Length == 0 && ext.Length > 1)
            {
                name = ext;
                ext = "";
            }
        }

        /// <summary>普通批量改名：替换 -> 大小写 -> 前缀 -> 后缀</summary>
        public static string ApplyBatchRules(string fileName, string find, string repl, bool useRegex,
                                             string caseMode, string addPrefix, string addSuffix)
        {
            string name;
            string ext;
            SplitName(fileName, out name, out ext);
            if (!string.IsNullOrEmpty(find))
            {
                if (useRegex)
                {
                    try
                    {
                        // Python 用 \1 做反向引用，.NET 用 $1；这里把 \1 归一化成 $1，避免静默产出错名
                        name = Regex.Replace(name, find, NormalizeRegexReplacement(repl ?? ""));
                    }
                    catch (ArgumentException ex) { throw new ArgumentException("正则表达式错误：" + ex.Message); }
                }
                else name = name.Replace(find, repl ?? "");
            }
            if (caseMode == "全部大写") name = name.ToUpperInvariant();
            else if (caseMode == "全部小写") name = name.ToLowerInvariant();
            else if (caseMode == "首字母大写")
            {
                TextInfo ti = CultureInfo.CurrentCulture.TextInfo;
                name = ti.ToTitleCase(name.ToLowerInvariant());
            }
            name = (addPrefix ?? "") + name + (addSuffix ?? "");
            return name + ext;
        }

        /// <summary>把 Python 风格的 \1 \2 反向引用转成 .NET 的 $1 $2；字面量 $ 转义成 $$</summary>
        public static string NormalizeRegexReplacement(string replacement)
        {
            if (string.IsNullOrEmpty(replacement)) return "";
            if (replacement.IndexOf('\\') < 0) return replacement;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < replacement.Length; i++)
            {
                char ch = replacement[i];
                if (ch == '\\' && i + 1 < replacement.Length && char.IsDigit(replacement[i + 1]))
                {
                    int j = i + 1;
                    while (j < replacement.Length && char.IsDigit(replacement[j])) j++;
                    sb.Append('$').Append(replacement, i + 1, j - i - 1);
                    i = j - 1;
                }
                else if (ch == '\\' && i + 1 < replacement.Length)
                {
                    // \n \t 之类保持原样交给 .NET
                    sb.Append(ch).Append(replacement[i + 1]);
                    i++;
                }
                else sb.Append(ch);
            }
            return sb.ToString();
        }

        /// <summary>通配符匹配，支持 * ? 和 [0-9] / [abc] 字符类（与 Python fnmatch 对齐）</summary>
        public static bool GlobMatch(string fileName, string pattern)
        {
            if (string.IsNullOrEmpty(pattern)) pattern = "*.*";
            StringBuilder sb = new StringBuilder("^");
            int i = 0;
            while (i < pattern.Length)
            {
                char ch = pattern[i];
                if (ch == '*') { sb.Append(".*"); i++; }
                else if (ch == '?') { sb.Append('.'); i++; }
                else if (ch == '[')
                {
                    int close = pattern.IndexOf(']', i + 1);
                    if (close > i + 1)
                    {
                        string inner = pattern.Substring(i + 1, close - i - 1);
                        sb.Append('[');
                        if (inner.StartsWith("!")) { sb.Append('^'); inner = inner.Substring(1); }
                        sb.Append(inner.Replace("\\", "\\\\"));
                        sb.Append(']');
                        i = close + 1;
                    }
                    else { sb.Append("\\["); i++; }
                }
                else { sb.Append(Regex.Escape(ch.ToString())); i++; }
            }
            sb.Append("$");
            try
            {
                return Regex.IsMatch(fileName, sb.ToString(), RegexOptions.IgnoreCase);
            }
            catch (ArgumentException)
            {
                return false;
            }
        }
    }

    // ---------------- 参数对象 ----------------

    public class ExcelRenameOptions
    {
        public string ExcelPath = "";
        public string Folder = "";
        public string ColPrefix = "";
        public string ColOldId = "";
        public string Separator = "_";
        public string Extension = "pdf";
        public int PadDigits = 0;
        public bool AutoClean = true;
        public string DupMode = "skip";    // skip | auto
        public string OpMode = "rename";   // rename | copy
        public string OutputFolder = "";
        public int StartRow = 2;           // Excel 行号（含表头行）
        public int EndRow = 0;             // 0 = 到最后一行
    }

    public class BatchRenameOptions
    {
        public string Folder = "";
        public string FileFilter = "*.pdf";
        public string AddPrefix = "";
        public string AddSuffix = "";
        public string FindText = "";
        public string ReplaceText = "";
        public bool UseRegex = false;
        public string CaseMode = "不变";
        public string OpMode = "rename";   // rename | copy
        public string OutputFolder = "";
        public string DupMode = "skip";    // skip | auto
    }

    // ---------------- 引擎 ----------------

    public static class RenameEngine
    {
        public static PreviewResult PreviewExcel(ExcelTable table, ExcelRenameOptions opt,
                                                 Action<LogItem> log = null, Action<int, int> progress = null)
        {
            PreviewResult result = new PreviewResult();
            result.Sink = log;
            if (table == null) throw new ArgumentException("请先加载 Excel");
            if (string.IsNullOrEmpty(opt.Folder) || !Directory.Exists(opt.Folder))
                throw new DirectoryNotFoundException("请选择文件所在文件夹");
            if (string.IsNullOrEmpty(opt.ColPrefix) || string.IsNullOrEmpty(opt.ColOldId))
                throw new ArgumentException("请选择前缀列和旧名字列");
            if (table.ColumnIndex(opt.ColPrefix) < 0)
                throw new ArgumentException("Excel 里找不到前缀列「" + opt.ColPrefix + "」，请重新加载 Excel 列。");
            if (table.ColumnIndex(opt.ColOldId) < 0)
                throw new ArgumentException("Excel 里找不到旧名字列「" + opt.ColOldId + "」，请重新加载 Excel 列。");

            int total = table.Rows.Count;
            if (total == 0)
            {
                result.AddLog("Excel 中没有数据行。", "miss");
                return result;
            }

            int dataStartRow = table.HeaderRow + 1;
            int startIdx = opt.StartRow - dataStartRow;
            if (startIdx < 0) startIdx = 0;
            int endIdx = opt.EndRow > table.HeaderRow ? opt.EndRow - table.HeaderRow : total;
            if (endIdx > total) endIdx = total;
            if (startIdx >= endIdx)
                throw new ArgumentException("行范围无效：起始行必须小于结束行");

            result.AddLog(string.Format("处理范围：Excel 第{0}~{1}行，共{2}条",
                startIdx + dataStartRow, endIdx + table.HeaderRow, endIdx - startIdx), "info");

            Dictionary<string, bool> seen = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            int cntOk = 0, cntMiss = 0, cntDup = 0;

            for (int i = startIdx; i < endIdx; i++)
            {
                if (progress != null) progress(i - startIdx + 1, endIdx - startIdx);

                string prefixRaw = table.Get(i, opt.ColPrefix);
                string oldRaw = table.Get(i, opt.ColOldId);
                if (string.IsNullOrEmpty(prefixRaw.Trim()) || string.IsNullOrEmpty(oldRaw.Trim())) continue;

                string oldFileName = NameRules.OldFileName(oldRaw, opt.Extension, opt.AutoClean);
                string oldFull = Path.Combine(opt.Folder, oldFileName);
                string newName = NameRules.BuildNewName(prefixRaw, oldRaw, opt.Separator, opt.Extension,
                                                        opt.PadDigits, opt.AutoClean);

                if (seen.ContainsKey(newName))
                {
                    if (opt.DupMode == "auto")
                    {
                        string b = Path.GetFileNameWithoutExtension(newName);
                        string e = Path.GetExtension(newName);
                        int n = 1;
                        while (seen.ContainsKey(b + "(" + n + ")" + e)) n++;
                        newName = b + "(" + n + ")" + e;
                        result.AddLog("[重名自动改名] " + oldFileName + " -> " + newName, "dup");
                    }
                    else
                    {
                        result.AddLog("[重名警告] 新文件名重复，将跳过：" + newName, "dup");
                        cntDup++;
                        continue;
                    }
                }
                seen[newName] = true;

                result.Items.Add(new RenamePair(oldFull, oldFileName, newName));

                if (File.Exists(oldFull))
                {
                    result.AddLog("[OK] " + oldFileName + " -> " + newName, "ok");
                    cntOk++;
                }
                else
                {
                    result.AddLog("[缺失] 找不到文件：" + oldFileName, "miss");
                    cntMiss++;
                }
            }

            result.OkCount = cntOk;
            result.MissingCount = cntMiss;
            result.DupCount = cntDup;
            result.AddLog(string.Format("预览统计：可改名 {0} 个，找不到文件 {1} 个，重名跳过 {2} 个", cntOk, cntMiss, cntDup),
                          (cntMiss > 0 || cntDup > 0) ? "warn" : "info");
            if (cntMiss > 0 || cntDup > 0)
                result.AddLog("⚠ 存在缺失或重名文件，执行时会再次提示确认", "warn");
            if (cntOk == 0) result.AddLog("✖ 没有任何可改名的文件，请检查文件夹、扩展名和 Excel 内容。", "miss");
            return result;
        }

        public static PreviewResult PreviewBatch(BatchRenameOptions opt, Action<LogItem> log = null,
                                                 Action<int, int> progress = null)
        {
            PreviewResult result = new PreviewResult();
            result.Sink = log;
            if (string.IsNullOrEmpty(opt.Folder) || !Directory.Exists(opt.Folder))
                throw new DirectoryNotFoundException("请先选择文件夹");

            string[] patterns = SplitPatterns(opt.FileFilter);
            string[] files = Directory.GetFiles(opt.Folder);
            List<string> matched = new List<string>();
            foreach (string f in files)
            {
                string name = Path.GetFileName(f);
                foreach (string p in patterns)
                {
                    if (NameRules.GlobMatch(name, p)) { matched.Add(name); break; }
                }
            }
            matched.Sort(StringComparer.OrdinalIgnoreCase);
            if (matched.Count == 0)
            {
                result.AddLog("没有找到匹配的文件", "miss");
                return result;
            }

            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int cntOk = 0, cntSkip = 0, cntDup = 0;
            int current = 0;
            foreach (string oldName in matched)
            {
                current++;
                if (progress != null) progress(current, matched.Count);
                string newName;
                try
                {
                    newName = NameRules.ApplyBatchRules(oldName, opt.FindText, opt.ReplaceText, opt.UseRegex,
                                                        opt.CaseMode, opt.AddPrefix, opt.AddSuffix);
                }
                catch (ArgumentException ex)
                {
                    result.AddLog("[错误] " + oldName + "：" + ex.Message, "miss");
                    cntSkip++;
                    continue;
                }
                if (newName == oldName)
                {
                    result.AddLog("[跳过] 名称未变：" + oldName, "info");
                    cntSkip++;
                    continue;
                }
                if (seen.Contains(newName))
                {
                    if (opt.DupMode == "auto")
                    {
                        string b = Path.GetFileNameWithoutExtension(newName);
                        string e = Path.GetExtension(newName);
                        int n = 1;
                        while (seen.Contains(b + "(" + n + ")" + e)) n++;
                        newName = b + "(" + n + ")" + e;
                    }
                    else
                    {
                        result.AddLog("[重名跳过] " + oldName + " -> " + newName, "dup");
                        cntSkip++;
                        cntDup++;
                        continue;
                    }
                }
                seen.Add(newName);
                result.Items.Add(new RenamePair(Path.Combine(opt.Folder, oldName), oldName, newName));
                result.AddLog("[OK] " + oldName + " -> " + newName, "ok");
                cntOk++;
            }

            result.OkCount = cntOk;
            result.DupCount = cntDup;
            result.AddLog(string.Format("预览统计：将处理 {0} 个，跳过 {1} 个（其中重名 {2} 个）", cntOk, cntSkip, cntDup), "info");
            if (cntOk == 0) result.AddLog("✖ 没有任何可处理的文件，请检查文件类型过滤和改名规则。", "miss");
            return result;
        }

        private static string[] SplitPatterns(string filter)
        {
            if (string.IsNullOrEmpty(filter)) return new string[] { "*.*" };
            string[] parts = filter.Split(new char[] { ' ', '\t', ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return new string[] { "*.*" };
            return parts;
        }

        public static ExecuteResult ExecuteExcel(IList<RenamePair> items, ExcelRenameOptions opt,
                                                 Action<LogItem> log = null, Action<int, int> progress = null)
        {
            return Execute(items, opt.OpMode, opt.OutputFolder, opt.Folder, opt.DupMode, "rename_record_",
                           log, progress);
        }

        public static ExecuteResult ExecuteBatch(IList<RenamePair> items, BatchRenameOptions opt,
                                                 Action<LogItem> log = null, Action<int, int> progress = null)
        {
            return Execute(items, opt.OpMode, opt.OutputFolder, opt.Folder, opt.DupMode, "batch_record_",
                           log, progress);
        }

        private static ExecuteResult Execute(IList<RenamePair> items, string opMode, string outputFolder,
                                             string sourceFolder, string dupMode, string recordPrefix,
                                             Action<LogItem> log, Action<int, int> progress)
        {
            ExecuteResult result = new ExecuteResult();
            result.Sink = log;
            if (items == null || items.Count == 0) throw new ArgumentException("没有可执行的项目，请先预览");

            bool isCopy = opMode == "copy";
            string targetFolder = sourceFolder;
            if (isCopy)
            {
                if (string.IsNullOrEmpty(outputFolder) || !Directory.Exists(outputFolder))
                    throw new DirectoryNotFoundException("复制模式下请先选择输出文件夹");
                if (SamePath(outputFolder, sourceFolder))
                    throw new ArgumentException("输出文件夹不能和原文件夹相同");
                targetFolder = outputFolder;
            }

            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string recordFile = Path.Combine(targetFolder, recordPrefix + stamp + ".csv");
            List<RenamePair> records = new List<RenamePair>();
            int cntSuccess = 0, cntFail = 0;
            int i = 0;

            foreach (RenamePair item in items)
            {
                i++;
                if (progress != null) progress(i, items.Count);

                if (!File.Exists(item.OldFull))
                {
                    result.AddLog("跳过：文件不存在 " + item.OldName);
                    cntFail++;
                    continue;
                }
                string newName = item.NewName;
                string newFull = Path.Combine(targetFolder, newName);

                // Windows 文件名大小写不敏感：report.pdf -> Report.pdf 会被 File.Exists 当成“已存在”，
                // 实际是同一个文件，必须放行，否则纯大小写改名永远失败。
                bool caseOnlyRename = !isCopy
                    && string.Equals(newFull, item.OldFull, StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(newFull, item.OldFull, StringComparison.Ordinal);

                if (!caseOnlyRename && File.Exists(newFull))
                {
                    if (dupMode == "auto")
                    {
                        string b = Path.Combine(targetFolder, Path.GetFileNameWithoutExtension(newName));
                        string e = Path.GetExtension(newName);
                        int n = 1;
                        while (File.Exists(b + "(" + n + ")" + e)) n++;
                        newFull = b + "(" + n + ")" + e;
                        newName = Path.GetFileName(newFull);
                    }
                    else
                    {
                        result.AddLog("跳过：目标文件名已存在 " + newName);
                        cntFail++;
                        continue;
                    }
                }
                try
                {
                    if (isCopy)
                    {
                        File.Copy(item.OldFull, newFull, false);
                        // 源文件只读时副本也只读，SetLastWriteTime 会抛异常，
                        // 造成“其实复制成功却报失败、还没写记录”。先去掉只读位再保留时间戳。
                        try
                        {
                            FileInfo fi = new FileInfo(newFull);
                            if ((fi.Attributes & FileAttributes.ReadOnly) != 0)
                                fi.Attributes = fi.Attributes & ~FileAttributes.ReadOnly;
                            fi.LastWriteTime = File.GetLastWriteTime(item.OldFull);
                        }
                        catch (Exception timeEx)
                        {
                            result.AddLog("提示：" + newName + " 已复制，但未能保留修改时间（" + timeEx.Message + "）");
                        }
                        result.AddLog("已复制：" + item.OldName + " -> " + newName);
                    }
                    else
                    {
                        File.Move(item.OldFull, newFull);
                        result.AddLog("已改名：" + item.OldName + " -> " + newName);
                    }
                    records.Add(new RenamePair(item.OldFull, item.OldName, newName));
                    cntSuccess++;
                }
                catch (Exception ex)
                {
                    result.AddLog((isCopy ? "复制失败 " : "改名失败 ") + item.OldName + " : " + ex.Message);
                    cntFail++;
                }
            }

            if (records.Count > 0)
            {
                try
                {
                    WriteRecords(recordFile, records);
                    result.RecordFile = recordFile;
                    result.AddLog("改名记录已保存：" + recordFile, "info");
                }
                catch (Exception ex)
                {
                    result.AddLog("保存记录失败：" + ex.Message, "miss");
                }
            }

            result.Success = cntSuccess;
            result.Failed = cntFail;
            return result;
        }

        private static void WriteRecords(string file, List<RenamePair> records)
        {
            using (StreamWriter w = new StreamWriter(file, false, new UTF8Encoding(true)))
            {
                w.WriteLine("old_name,new_name");
                foreach (RenamePair r in records)
                {
                    w.WriteLine(CsvEscape(r.OldName) + "," + CsvEscape(r.NewName));
                }
            }
        }

        private static string CsvEscape(string s)
        {
            if (s == null) return "";
            if (s.IndexOf(',') >= 0 || s.IndexOf('"') >= 0 || s.IndexOf('\n') >= 0 || s.IndexOf('\r') >= 0)
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        private static List<RenamePair> ReadRecords(string file)
        {
            List<RenamePair> list = new List<RenamePair>();
            string text = File.ReadAllText(file, Encoding.UTF8);
            string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            bool header = true;
            foreach (string line in lines)
            {
                if (line.Length == 0) continue;
                if (header) { header = false; continue; }
                string[] cells = SplitCsvLine(line);
                if (cells.Length < 2) continue;
                string oldName = cells[0].Trim();
                string newName = cells[1].Trim();
                if (oldName.Length == 0 || newName.Length == 0) continue;
                list.Add(new RenamePair("", oldName, newName));
            }
            return list;
        }

        private static string[] SplitCsvLine(string line)
        {
            List<string> cells = new List<string>();
            StringBuilder cur = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];
                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"') { cur.Append('"'); i++; }
                        else inQuotes = false;
                    }
                    else cur.Append(ch);
                }
                else
                {
                    if (ch == '"') inQuotes = true;
                    else if (ch == ',') { cells.Add(cur.ToString()); cur.Length = 0; }
                    else cur.Append(ch);
                }
            }
            cells.Add(cur.ToString());
            return cells.ToArray();
        }

        /// <summary>撤销：sourceFolder 是界面上选择的原文件夹</summary>
        public static ExecuteResult Undo(string recordFile, string sourceFolder,
                                         Action<LogItem> log = null, Action<int, int> progress = null)
        {
            ExecuteResult result = new ExecuteResult();
            result.Sink = log;
            if (string.IsNullOrEmpty(recordFile) || !File.Exists(recordFile))
                throw new FileNotFoundException("记录文件不存在：" + recordFile);
            if (string.IsNullOrEmpty(sourceFolder) || !Directory.Exists(sourceFolder))
                throw new DirectoryNotFoundException("请先选择文件所在文件夹");

            string recordDir = Path.GetDirectoryName(Path.GetFullPath(recordFile));
            bool isCopyMode = !SamePath(recordDir, sourceFolder);
            string targetDir = isCopyMode ? recordDir : sourceFolder;

            List<RenamePair> records = ReadRecords(recordFile);
            int cntOk = 0, cntFail = 0;
            int i = 0;
            foreach (RenamePair rec in records)
            {
                i++;
                if (progress != null) progress(i, records.Count);
                string newFull = Path.Combine(targetDir, rec.NewName);
                string oldFull = Path.Combine(targetDir, rec.OldName);

                if (!File.Exists(newFull))
                {
                    result.AddLog("跳过：文件不存在 " + rec.NewName);
                    cntFail++;
                    continue;
                }
                if (isCopyMode)
                {
                    try
                    {
                        File.Delete(newFull);
                        result.AddLog("已删除副本：" + rec.NewName);
                        cntOk++;
                    }
                    catch (Exception ex)
                    {
                        result.AddLog("删除失败 " + rec.NewName + " : " + ex.Message);
                        cntFail++;
                    }
                }
                else
                {
                    bool caseOnly = string.Equals(newFull, oldFull, StringComparison.OrdinalIgnoreCase)
                                    && !string.Equals(newFull, oldFull, StringComparison.Ordinal);
                    if (!caseOnly && File.Exists(oldFull))
                    {
                        result.AddLog("跳过：原名文件已存在 " + rec.OldName);
                        cntFail++;
                        continue;
                    }
                    try
                    {
                        File.Move(newFull, oldFull);
                        result.AddLog("已撤销：" + rec.NewName + " -> " + rec.OldName);
                        cntOk++;
                    }
                    catch (Exception ex)
                    {
                        result.AddLog("撤销失败 " + rec.NewName + " : " + ex.Message);
                        cntFail++;
                    }
                }
            }
            result.Success = cntOk;
            result.Failed = cntFail;
            return result;
        }

        public static bool SamePath(string a, string b)
        {
            if (a == null || b == null) return false;
            return string.Equals(Path.GetFullPath(a).TrimEnd('\\'), Path.GetFullPath(b).TrimEnd('\\'),
                                 StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>判断记录文件是复制模式生成的（记录目录 != 源文件夹）</summary>
        public static bool IsCopyModeRecord(string recordFile, string sourceFolder)
        {
            string recordDir = Path.GetDirectoryName(Path.GetFullPath(recordFile));
            return !SamePath(recordDir, sourceFolder);
        }
    }
}