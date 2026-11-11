using System;
using System.IO;

namespace ExcelBatchRename
{
    internal static class EdgeTest
    {
        private static int _fail;
        private static void Check(bool c, string m)
        {
            Console.WriteLine((c ? "[PASS] " : "[FAIL] ") + m);
            if (!c) _fail++;
        }

        public static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            string root = Path.Combine(Directory.GetCurrentDirectory(), "testdata", "边界测试");
            string src = Path.Combine(root, "文件");
            ExcelTable t = ExcelTable.Load(Path.Combine(root, "边界.xlsx"));
            Console.WriteLine("sheet=" + t.SheetName + " headers=" + string.Join("|", t.Headers) + " rows=" + t.Rows.Count);
            Check(t.SheetName == "数据", "只读第一个工作表");
            Check(t.Headers.Count == 2 && t.Headers[0] == "编号", "表头正确");
            Check(t.Rows.Count == 6, "空行占位后行数=6，实际 " + t.Rows.Count);
            Check(t.Get(0, "编号") == "1", "1.0 -> \"1\"，实际 " + t.Get(0, "编号"));
            Check(t.Get(1, "文件编号") == "B 002.pdf", "带扩展名的单元格");
            Check(t.Get(2, "编号") == "", "空行读为空");
            Check(t.Get(3, "文件编号") == "C 003", "空行之后的数据行");
            Check(t.Get(5, "编号") == "注释行", "单列行 + 缺列补空");

            ExcelRenameOptions o = new ExcelRenameOptions();
            o.Folder = src; o.ColPrefix = "编号"; o.ColOldId = "文件编号";
            o.Separator = "_"; o.Extension = "pdf"; o.PadDigits = 3;
            o.AutoClean = true; o.DupMode = "skip"; o.OpMode = "rename";
            o.StartRow = 2; o.EndRow = 0;
            PreviewResult pv = RenameEngine.PreviewExcel(t, o);
            foreach (LogItem li in pv.Log) Console.WriteLine("   " + li.Text);
            Check(pv.Items.Count == 4, "预览拿到 4 条有效行，实际 " + pv.Items.Count);
            if (pv.Items.Count == 4)
            {
                Check(pv.Items[0].NewName == "001_A 001.pdf", "第 1 条：" + pv.Items[0].NewName);
                Check(pv.Items[1].NewName == "002_B 002.pdf", "第 2 条去掉重复扩展名：" + pv.Items[1].NewName);
            }
            Check(pv.MissingCount == 0, "4 个文件都找得到，缺失 " + pv.MissingCount);

            ExcelRenameOptions o2 = new ExcelRenameOptions();
            o2.Folder = src; o2.ColPrefix = "编号"; o2.ColOldId = "文件编号";
            o2.Separator = "_"; o2.Extension = ".PDF"; o2.PadDigits = 0;
            o2.AutoClean = true; o2.DupMode = "skip"; o2.OpMode = "rename";
            o2.StartRow = 2; o2.EndRow = 0;
            PreviewResult pv2 = RenameEngine.PreviewExcel(t, o2);
            Check(pv2.Items.Count == 4 && pv2.Items[0].NewName == "1_A 001.PDF",
                  "扩展名写 .PDF 也能用：" + (pv2.Items.Count > 0 ? pv2.Items[0].NewName : "<无>"));

            string bad = Path.Combine(root, "假名.xls");
            File.WriteAllText(bad, "x");
            try { ExcelTable.Load(bad); Check(false, "xls 应该抛异常"); }
            catch (NotSupportedException ex) { Check(true, "xls 给出明确提示"); }

            try { ExcelTable.Load(Path.Combine(root, "不存在.xlsx")); Check(false, "缺文件应抛异常"); }
            catch (FileNotFoundException) { Check(true, "缺文件给出明确提示"); }

            try
            {
                o2.StartRow = 9; o2.EndRow = 3;
                RenameEngine.PreviewExcel(t, o2);
                Check(false, "非法行范围应该抛异常");
            }
            catch (ArgumentException) { Check(true, "非法行范围被拒绝"); }

            // Excel 列名不存在
            ExcelRenameOptions o3 = new ExcelRenameOptions();
            o3.Folder = src; o3.ColPrefix = "不存在的列"; o3.ColOldId = "文件编号";
            o3.Extension = "pdf"; o3.StartRow = 2; o3.EndRow = 0; o3.AutoClean = true;
            try
            {
                RenameEngine.PreviewExcel(t, o3);
                Check(false, "列名不存在应该抛异常");
            }
            catch (ArgumentException ex2)
            {
                Check(ex2.Message.Contains("找不到前缀列"), "列名不存在时给出明确提示：" + ex2.Message);
            }

            Console.WriteLine(_fail == 0 ? "== 边界测试全部通过 ==" : "== 有 " + _fail + " 项失败 ==");
            return _fail == 0 ? 0 : 1;
        }
    }
}