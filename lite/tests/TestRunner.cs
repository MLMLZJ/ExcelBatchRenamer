using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelBatchRename
{
    internal static class TestRunner
    {
        private static int _fail;

        private static void Check(bool cond, string what)
        {
            Console.WriteLine((cond ? "[PASS] " : "[FAIL] ") + what);
            if (!cond) _fail++;
        }

        private static string Root = Path.Combine(Directory.GetCurrentDirectory(), "testdata");

        public static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            string root = Root;
            string src = Path.Combine(root, "源文件夹");
            string outDir = Path.Combine(root, "输出文件夹");
            string excel = Path.Combine(root, "清单.xlsx");

            // ---------- 1. Excel 读取 ----------
            ExcelTable table = ExcelTable.Load(excel);
            Console.WriteLine("工作表: " + table.SheetName + "  表头: " + string.Join(" | ", table.Headers.ToArray()));
            Check(table.SheetName == "清单", "工作表名读取正确");
            Check(table.Headers.Count == 2 && table.Headers[0] == "编号" && table.Headers[1] == "文件编号", "表头读取正确");
            Check(table.Rows.Count == 9, "数据行数 = 9，实际 " + table.Rows.Count);
            Check(table.Get(4, "文件编号") == "FILE10005.pdf", "文本单元格读取正确");
            Check(table.Get(0, "编号") == "1", "数字单元格 1 被规范成 \"1\"，实际 " + table.Get(0, "编号"));
            Check(table.Get(8, "编号") == "A08", "文本前缀 A08 读取正确");

            // ---------- 2. 命名规则 ----------
            Check(NameRules.Clean("  a/b:c  ") == "a_b_c", "清洗文件名");
            Check(NameRules.StripExt("FILE10001.PDF", "pdf") == "FILE10001", "去扩展名（大小写无关）");
            Check(NameRules.PadPrefix("3", 3) == "003", "补零 3 -> 003");
            Check(NameRules.PadPrefix("A1", 3) == "A1", "非数字前缀不补零");
            Check(NameRules.BuildNewName("1", "FILE10001", "_", "pdf", 3, true) == "001_FILE10001.pdf", "构造新文件名");
            Check(NameRules.OldFileName("FILE10001.pdf", "pdf", true) == "FILE10001.pdf", "旧名字带扩展名去重");

            // ---------- 3. Excel 预览 ----------
            ExcelRenameOptions opt = new ExcelRenameOptions();
            opt.ExcelPath = excel;
            opt.Folder = src;
            opt.ColPrefix = "编号";
            opt.ColOldId = "文件编号";
            opt.Separator = "_";
            opt.Extension = "pdf";
            opt.PadDigits = 3;
            opt.AutoClean = true;
            opt.DupMode = "skip";
            opt.OpMode = "copy";
            opt.OutputFolder = outDir;
            opt.StartRow = 2;
            opt.EndRow = 0;

            PreviewResult pv = RenameEngine.PreviewExcel(table, opt);
            foreach (LogItem li in pv.Log) Console.WriteLine("   " + li.Text);
            Check(pv.Items.Count == 8, "预览生成 8 条（1 条重名跳过），实际 " + pv.Items.Count);
            Check(pv.MissingCount == 1, "缺失 1 个，实际 " + pv.MissingCount);
            Check(pv.DupCount == 1, "重名 1 个，实际 " + pv.DupCount);
            Check(pv.Items[0].NewName == "001_FILE10001.pdf", "第 1 条新名字正确");

            // 重名 auto 模式
            ExcelRenameOptions autoOpt = new ExcelRenameOptions();
            autoOpt.Folder = src; autoOpt.ColPrefix = "编号"; autoOpt.ColOldId = "文件编号";
            autoOpt.Separator = "_"; autoOpt.Extension = "pdf"; autoOpt.PadDigits = 0;
            autoOpt.AutoClean = true; autoOpt.DupMode = "auto"; autoOpt.OpMode = "rename";
            autoOpt.StartRow = 2; autoOpt.EndRow = 0;
            PreviewResult pvAuto = RenameEngine.PreviewExcel(table, autoOpt);
            Check(pvAuto.Items.Count == 9, "自动加序号时 9 条全保留，实际 " + pvAuto.Items.Count);
            Check(pvAuto.Items[7].NewName == "1_FILE10001(1).pdf", "第 8 条自动编号，" + pvAuto.Items[7].NewName);

            // 行范围
            ExcelRenameOptions rangeOpt = new ExcelRenameOptions();
            rangeOpt.Folder = src; rangeOpt.ColPrefix = "编号"; rangeOpt.ColOldId = "文件编号";
            rangeOpt.Separator = "_"; rangeOpt.Extension = "pdf"; rangeOpt.PadDigits = 3;
            rangeOpt.AutoClean = true; rangeOpt.DupMode = "skip"; rangeOpt.OpMode = "rename";
            rangeOpt.StartRow = 3; rangeOpt.EndRow = 5;
            PreviewResult pvRange = RenameEngine.PreviewExcel(table, rangeOpt);
            Check(pvRange.Items.Count == 3, "行范围 3~5 命中 3 条，实际 " + pvRange.Items.Count);
            Check(pvRange.Items[0].NewName == "002_FILE10002.pdf", "行范围首条正确");

            // ---------- 4. 复制模式执行 ----------
            ExecuteResult er = RenameEngine.ExecuteExcel(pv.Items, opt);
            foreach (LogItem li in er.Log) Console.WriteLine("   " + li.Text);
            Check(er.Success == 7, "复制成功 7 个（第 8 条源文件缺失），实际 " + er.Success);
            Check(er.Failed == 1, "失败 1 个，实际 " + er.Failed);
            Check(File.Exists(Path.Combine(outDir, "001_FILE10001.pdf")), "输出文件存在");
            Check(File.Exists(Path.Combine(src, "FILE10001.pdf")), "原文件仍然存在（复制模式）");
            Check(er.RecordFile != null && File.Exists(er.RecordFile), "记录文件已生成");

            // ---------- 5. 撤销复制 ----------
            ExecuteResult undo = RenameEngine.Undo(er.RecordFile, src);
            foreach (LogItem li in undo.Log) Console.WriteLine("   " + li.Text);
            Check(undo.Success == 7, "撤销副本 7 个，实际 " + undo.Success);
            Check(!File.Exists(Path.Combine(outDir, "001_FILE10001.pdf")), "输出副本已删除");

            // ---------- 6. 直接改名 + 撤销 ----------
            string renameDir = Path.Combine(root, "改名测试");
            if (Directory.Exists(renameDir)) Directory.Delete(renameDir, true);
            Directory.CreateDirectory(renameDir);
            foreach (string n in new string[] { "FILE10001.pdf", "FILE10002.pdf", "FILE10003.pdf" })
                File.WriteAllText(Path.Combine(renameDir, n), "x");

            ExcelRenameOptions rn = new ExcelRenameOptions();
            rn.Folder = renameDir; rn.ColPrefix = "编号"; rn.ColOldId = "文件编号";
            rn.Separator = "-"; rn.Extension = "pdf"; rn.PadDigits = 0;
            rn.AutoClean = true; rn.DupMode = "skip"; rn.OpMode = "rename";
            rn.StartRow = 2; rn.EndRow = 4;
            PreviewResult pvRename = RenameEngine.PreviewExcel(table, rn);
            ExecuteResult er2 = RenameEngine.ExecuteExcel(pvRename.Items, rn);
            Check(er2.Success == 3, "直接改名 3 个，实际 " + er2.Success);
            Check(File.Exists(Path.Combine(renameDir, "1-FILE10001.pdf")), "改名结果正确");
            Check(!File.Exists(Path.Combine(renameDir, "FILE10001.pdf")), "原文件已不存在");
            ExecuteResult undo2 = RenameEngine.Undo(er2.RecordFile, renameDir);
            Check(undo2.Success == 3, "撤销改名 3 个，实际 " + undo2.Success);
            Check(File.Exists(Path.Combine(renameDir, "FILE10001.pdf")), "原名已恢复");

            // ---------- 7. 普通批量改名 ----------
            string batchDir = Path.Combine(root, "批量测试");
            BatchRenameOptions bopt = new BatchRenameOptions();
            bopt.Folder = batchDir;
            bopt.FileFilter = "*.pdf";
            bopt.AddPrefix = "2024_";
            bopt.AddSuffix = "_v2";
            bopt.FindText = "final";
            bopt.ReplaceText = "最终";
            bopt.UseRegex = false;
            bopt.CaseMode = "不变";
            bopt.OpMode = "copy";
            bopt.OutputFolder = Path.Combine(root, "批量输出");
            bopt.DupMode = "skip";
            if (Directory.Exists(bopt.OutputFolder)) Directory.Delete(bopt.OutputFolder, true);
            Directory.CreateDirectory(bopt.OutputFolder);

            PreviewResult bpv = RenameEngine.PreviewBatch(bopt);
            foreach (LogItem li in bpv.Log) Console.WriteLine("   " + li.Text);
            Check(bpv.Items.Count == 2, "批量预览命中 2 个 pdf，实际 " + bpv.Items.Count);
            Check(bpv.Items[0].NewName == "2024_最终 report 2_v2.pdf", "批量新名字正确：" + bpv.Items[0].NewName);
            Check(bpv.Items[1].NewName == "2024_最终 report_v2.pdf", "批量第 2 条新名字正确：" + bpv.Items[1].NewName);

            ExecuteResult ber = RenameEngine.ExecuteBatch(bpv.Items, bopt);
            Check(ber.Success == 2, "批量复制 2 个，实际 " + ber.Success);
            Check(File.Exists(Path.Combine(bopt.OutputFolder, "2024_最终 report_v2.pdf")), "批量输出文件存在");
            ExecuteResult bundo = RenameEngine.Undo(ber.RecordFile, batchDir);
            Check(bundo.Success == 2, "批量撤销 2 个，实际 " + bundo.Success);

            // 正则替换
            BatchRenameOptions ropt = new BatchRenameOptions();
            ropt.Folder = batchDir; ropt.FileFilter = "*.JPG"; ropt.AddPrefix = ""; ropt.AddSuffix = "";
            ropt.FindText = "IMG_(\\d+)"; ropt.ReplaceText = "照片$1"; ropt.UseRegex = true;
            ropt.CaseMode = "全部小写"; ropt.OpMode = "copy"; ropt.OutputFolder = bopt.OutputFolder;
            ropt.DupMode = "skip";
            PreviewResult rpv = RenameEngine.PreviewBatch(ropt);
            Check(rpv.Items.Count == 1 && rpv.Items[0].NewName == "照片001.JPG", "正则替换 + 小写：" + (rpv.Items.Count>0?rpv.Items[0].NewName:"<none>"));

            Console.WriteLine();
            Console.WriteLine(_fail == 0 ? "== 全部测试通过 ==" : "== 有 " + _fail + " 项失败 ==");
            return _fail == 0 ? 0 : 1;
        }
    }
}