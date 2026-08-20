using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ExcelBatchRenamer.Models;

namespace ExcelBatchRenamer.Services;

/// <summary>
/// Excel读取服务
/// </summary>
public class ExcelReader
{
    /// <summary>
    /// 读取Excel文件，返回表头和所有数据行
    /// </summary>
    public (List<string> headers, List<string[]> rows) Read(string filePath)
    {
        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        IWorkbook workbook = new XSSFWorkbook(fs);
        ISheet sheet = workbook.GetSheetAt(0);

        if (sheet.LastRowNum < 0)
            return (new List<string>(), new List<string[]>());

        // 读取表头（第一行）
        var headers = new List<string>();
        IRow headerRow = sheet.GetRow(0);
        int colCount = headerRow?.LastCellNum ?? 0;
        for (int i = 0; i < colCount; i++)
        {
            string? val = headerRow?.GetCell(i)?.ToString();
            headers.Add(string.IsNullOrWhiteSpace(val) ? $"列{i + 1}" : val.Trim());
        }

        // 读取数据行
        var rows = new List<string[]>();
        for (int r = 1; r <= sheet.LastRowNum; r++)
        {
            IRow? row = sheet.GetRow(r);
            if (row == null) continue;

            // 跳过整行空行
            bool hasData = false;
            var data = new string[colCount];
            for (int c = 0; c < colCount; c++)
            {
                string? val = row.GetCell(c)?.ToString();
                data[c] = val ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(data[c]))
                    hasData = true;
            }
            if (hasData)
                rows.Add(data);
        }

        workbook.Close();
        return (headers, rows);
    }

    /// <summary>
    /// 从Excel数据中提取改名映射项
    /// </summary>
    /// <param name="headers">表头列表</param>
    /// <param name="rows">数据行</param>
    /// <param name="prefixColIndex">前缀列索引（0-based）</param>
    /// <param name="oldNameColIndex">旧名列索引（0-based）</param>
    /// <param name="startRow">起始Excel行号（含表头，从1开始）</param>
    /// <param name="endRow">结束Excel行号（含表头，0或负数表示到最后）</param>
    public List<RenameItem> ExtractItems(
        List<string> headers,
        List<string[]> rows,
        int prefixColIndex,
        int oldNameColIndex,
        int startRow = 2,
        int endRow = 0)
    {
        var items = new List<RenameItem>();
        // rows[0] 对应 Excel 第2行（行号2）
        int startIdx = Math.Max(0, startRow - 2);
        int endIdx = endRow > 1 ? Math.Min(rows.Count - 1, endRow - 2) : rows.Count - 1;

        for (int i = startIdx; i <= endIdx; i++)
        {
            if (i < 0 || i >= rows.Count) continue;
            var row = rows[i];
            string prefix = prefixColIndex < row.Length ? row[prefixColIndex].Trim() : string.Empty;
            string oldName = oldNameColIndex < row.Length ? row[oldNameColIndex].Trim() : string.Empty;

            if (string.IsNullOrWhiteSpace(oldName))
                continue;

            items.Add(new RenameItem
            {
                RowNumber = i + 2, // Excel行号（含表头）
                Prefix = prefix,
                OldFileName = oldName
            });
        }
        return items;
    }
}
