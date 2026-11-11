using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelBatchRename
{
    internal static class FixTest
    {
        private static int _fail;
        private static void Check(bool c, string m)
        {
            Console.WriteLine((c ? "[PASS] " : "[FAIL] ") + m);
            if (!c) _fail++;
        }
        private static string MakeTemp(string tag)
        {
            string p = Path.Combine(Path.GetTempPath(), "fix_" + tag + "_" + Guid.NewGuid().ToString("N").Substring(0, 6));
            Directory.CreateDirectory(p);
            return p;
        }

        public static int Main()
        {
            Console.OutputEncoding = Encoding.UTF8;

            // ---------- 1. 仅大小写改名 ----------
            string d1 = MakeTemp("case");
            File.WriteAllText(Path.Combine(d1, "report.pdf"), "x");
            BatchRenameOptions b1 = new BatchRenameOptions();
            b1.Folder = d1; b1.FileFilter = "*.pdf"; b1.CaseMode = "全部大写";
            b1.OpMode = "rename"; b1.DupMode = "skip";
            PreviewResult p1 = RenameEngine.PreviewBatch(b1);
            Check(p1.Items.Count == 1, "预览出 1 条大小写改名，实际 " + p1.Items.Count);
            ExecuteResult e1 = RenameEngine.ExecuteBatch(p1.Items, b1);
            Check(e1.Success == 1, "大小写改名成功 1 个（旧版 0），实际 " + e1.Success);
            Check(File.Exists(Path.Combine(d1, "REPORT.PDF")), "磁盘上出现 REPORT.PDF");
            DeleteDir(d1);

            // 大小写撤销
            string d1b = MakeTemp("caseundo");
            File.WriteAllText(Path.Combine(d1b, "abc.pdf"), "x");
            BatchRenameOptions bb = new BatchRenameOptions();
            bb.Folder = d1b; bb.FileFilter = "*.pdf"; bb.CaseMode = "全部大写";
            bb.OpMode = "rename"; bb.DupMode = "skip";
            ExecuteResult eb = RenameEngine.ExecuteBatch(RenameEngine.PreviewBatch(bb).Items, bb);
            ExecuteResult ub = RenameEngine.Undo(eb.RecordFile, d1b);
            Check(ub.Success == 1 && File.Exists(Path.Combine(d1b, "abc.pdf")), "大小写改名可撤销，成功 " + ub.Success);
            DeleteDir(d1b);

            // ---------- 2. 只读源文件 + 复制模式 ----------
            string d2 = MakeTemp("ro");
            string src2 = Path.Combine(d2, "src");
            string out2 = Path.Combine(d2, "out");
            Directory.CreateDirectory(src2);
            Directory.CreateDirectory(out2);
            string f2 = Path.Combine(src2, "ro.pdf");
            File.WriteAllText(f2, "x");
            File.SetAttributes(f2, FileAttributes.ReadOnly);
            BatchRenameOptions b2 = new BatchRenameOptions();
            b2.Folder = src2; b2.FileFilter = "*.pdf"; b2.AddPrefix = "new_";
            b2.OpMode = "copy"; b2.OutputFolder = out2; b2.DupMode = "skip";
            PreviewResult p2 = RenameEngine.PreviewBatch(b2);
            ExecuteResult e2 = RenameEngine.ExecuteBatch(p2.Items, b2);
            Check(e2.Success == 1, "只读源文件复制成功 1 个，实际 " + e2.Success);
            Check(e2.RecordFile != null && File.Exists(e2.RecordFile), "只读复制也写了记录 CSV");
            Check(File.Exists(Path.Combine(out2, "new_ro.pdf")), "输出文件存在");
            ExecuteResult u2 = RenameEngine.Undo(e2.RecordFile, src2);
            Check(u2.Success == 1, "只读副本可撤销删除，实际 " + u2.Success);
            DeleteDir(d2);

            // ---------- 3. 前缀补零负数 ----------
            Check(NameRules.PadPrefix("-1", 3) == "-01", "负数补零 -01，实际 " + NameRules.PadPrefix("-1", 3));
            Check(NameRules.PadPrefix("3", 3) == "003", "正数补零 003");

            // ---------- 4. rPh 注音不混入 ----------
            string rph = BuildRphXlsx();
            ExcelTable t4 = ExcelTable.Load(rph);
            Console.WriteLine("   rPh 单元格读出：[" + t4.Get(0, "文件编号") + "]");
            Check(t4.Get(0, "文件编号") == "FILE10001", "rPh 注音被忽略，读出 FILE10001");
            DeleteDir(Path.GetDirectoryName(rph));

            // ---------- 5. 重名表头 ----------
            string dup = BuildDupHeaderXlsx();
            ExcelTable t5 = ExcelTable.Load(dup);
            Console.WriteLine("   表头：" + string.Join(" | ", t5.Headers.ToArray()));
            Check(t5.Headers.Count == 3 && t5.Headers[1] == "名(2)", "重名表头加 (2) 后缀");
            Check(t5.ColumnIndex("名(2)") == 1, "第二列现在可被选中");
            DeleteDir(Path.GetDirectoryName(dup));

            // ---------- 6. 点文件 ----------
            string n, x;
            NameRules.SplitName(".gitignore", out n, out x);
            Check(n == ".gitignore" && x == "", "点文件整体当主名（与 Python splitext 一致）");
            NameRules.SplitName("a.b.pdf", out n, out x);
            Check(n == "a.b" && x == ".pdf", "普通文件拆分正常");

            // ---------- 7. 正则反向引用 \1 ----------
            string r7 = NameRules.ApplyBatchRules("IMG_001.jpg", "IMG_(\\d+)", "照片\\1", true, "不变", "", "");
            Check(r7 == "照片001.jpg", "Python 风格 \\1 反向引用，实际 " + r7);
            string r7b = NameRules.ApplyBatchRules("IMG_001.jpg", "IMG_(\\d+)", "照片$1", true, "不变", "", "");
            Check(r7b == "照片001.jpg", ".NET 风格 $1 仍可用，实际 " + r7b);

            // ---------- 8. 通配符字符类 ----------
            Check(NameRules.GlobMatch("file1.pdf", "file[0-9].pdf"), "[0-9] 字符类匹配");
            Check(!NameRules.GlobMatch("fileX.pdf", "file[0-9].pdf"), "[0-9] 字符类不匹配字母");

            // ---------- 9. 表头行号 ----------
            ExcelTable t9 = ExcelTable.Load(@"C:\deepseek\批量改名\testdata\边界测试\边界.xlsx");
            Check(t9.HeaderRow == 1, "表头行号 = 1，实际 " + t9.HeaderRow);

            Console.WriteLine(_fail == 0 ? "== 修复项全部验证通过 ==" : "== 有 " + _fail + " 项失败 ==");
            return _fail == 0 ? 0 : 1;
        }

        private static void DeleteDir(string dir)
        {
            try
            {
                foreach (string f in Directory.GetFiles(dir)) File.SetAttributes(f, FileAttributes.Normal);
                Directory.Delete(dir, true);
            }
            catch { }
        }

        /// <summary>造一个含 rPh（注音）的 xlsx</summary>
        private static string BuildRphXlsx()
        {
            string dir = Path.Combine(Path.GetTempPath(), "rph_" + Guid.NewGuid().ToString("N").Substring(0, 6));
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "rph.xlsx");
            using (FileStream fs = new FileStream(path, FileMode.Create))
            using (var zip = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Create))
            {
                Write(zip, "xl/workbook.xml",
                    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"S1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                Write(zip, "xl/_rels/workbook.xml.rels",
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
                Write(zip, "xl/sharedStrings.xml",
                    "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"3\">"
                    + "<si><t>编号</t></si>"
                    + "<si><t>文件编号</t></si>"
                    + "<si><t>FILE10001</t><rPh sb=\"0\" eb=\"4\"><t>ファイル</t></rPh><phoneticPr fontId=\"1\"/></si>"
                    + "</sst>");
                Write(zip, "xl/worksheets/sheet1.xml",
                    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                    + "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c></row>"
                    + "<row r=\"2\"><c r=\"A2\"><v>1</v></c><c r=\"B2\" t=\"s\"><v>2</v></c></row>"
                    + "</sheetData></worksheet>");
            }
            return path;
        }

        private static string BuildDupHeaderXlsx()
        {
            string dir = Path.Combine(Path.GetTempPath(), "dup_" + Guid.NewGuid().ToString("N").Substring(0, 6));
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "dup.xlsx");
            using (FileStream fs = new FileStream(path, FileMode.Create))
            using (var zip = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Create))
            {
                Write(zip, "xl/worksheets/sheet1.xml",
                    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                    + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>名</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>名</t></is></c><c r=\"C1\" t=\"inlineStr\"><is><t>号</t></is></c></row>"
                    + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>x</t></is></c><c r=\"B2\" t=\"inlineStr\"><is><t>y</t></is></c><c r=\"C2\" t=\"inlineStr\"><is><t>z</t></is></c></row>"
                    + "</sheetData></worksheet>");
            }
            return path;
        }

        private static void Write(System.IO.Compression.ZipArchive zip, string name, string content)
        {
            var e = zip.CreateEntry(name);
            using (Stream s = e.Open())
            {
                byte[] b = Encoding.UTF8.GetBytes(content);
                s.Write(b, 0, b.Length);
            }
        }
    }
}
