using System.Text;
using ExcelBatchRenamer.Models;

namespace ExcelBatchRenamer.Services;

/// <summary>
/// 撤销记录管理
/// </summary>
public class UndoManager
{
    /// <summary>
    /// 保存改名记录到CSV文件
    /// </summary>
    public string SaveRecord(string targetFolder, List<RenameResult> results, OperationMode mode)
    {
        string recordFile = Path.Combine(targetFolder, $"rename_record_{DateTime.Now:yyyyMMdd_HHmmss}.csv");

        var sb = new StringBuilder();
        sb.AppendLine("mode,row,old_path,new_path,old_name,new_name");

        foreach (var r in results.Where(x => x.Success && !x.IsPreview))
        {
            sb.AppendLine($"{mode},{r.RowNumber},{EscapeCsv(r.OldPath)},{EscapeCsv(r.NewPath)},{EscapeCsv(r.OldName)},{EscapeCsv(r.NewName)}");
        }

        File.WriteAllText(recordFile, sb.ToString(), Encoding.UTF8);
        return recordFile;
    }

    /// <summary>
    /// 执行撤销操作
    /// </summary>
    public (int success, int failed, string message) Undo(string recordFile)
    {
        if (!File.Exists(recordFile))
            return (0, 0, "记录文件不存在");

        var lines = File.ReadAllLines(recordFile, Encoding.UTF8).Skip(1).ToList();
        if (lines.Count == 0)
            return (0, 0, "记录文件为空");

        int success = 0, failed = 0;
        var errors = new List<string>();

        // 倒序撤销，避免重名冲突
        for (int i = lines.Count - 1; i >= 0; i--)
        {
            var parts = ParseCsvLine(lines[i]);
            if (parts.Length < 4) continue;

            string mode = parts[0];
            string oldPath = parts[2];
            string newPath = parts[3];

            try
            {
                if (mode == nameof(OperationMode.Copy))
                {
                    // 复制模式：删除输出副本
                    if (File.Exists(newPath))
                    {
                        File.Delete(newPath);
                        success++;
                    }
                    else
                    {
                        failed++;
                        errors.Add($"文件已不存在: {newPath}");
                    }
                }
                else
                {
                    // 改名模式：改回原名
                    if (File.Exists(newPath))
                    {
                        // 确保目标目录存在
                        string? dir = Path.GetDirectoryName(oldPath);
                        if (!string.IsNullOrEmpty(dir))
                            Directory.CreateDirectory(dir);
                        File.Move(newPath, oldPath);
                        success++;
                    }
                    else
                    {
                        failed++;
                        errors.Add($"文件已不存在: {newPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                failed++;
                errors.Add($"{newPath}: {ex.Message}");
            }
        }

        string msg = $"撤销完成：成功 {success} 条，失败 {failed} 条";
        if (errors.Count > 0)
            msg += "\n" + string.Join("\n", errors.Take(10));
        return (success, failed, msg);
    }

    private static string EscapeCsv(string value)
    {
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        return value;
    }

    private static string[] ParseCsvLine(string line)
    {
        var result = new List<string>();
        bool inQuotes = false;
        var current = new StringBuilder();

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++;
                    }
                    else
                        inQuotes = false;
                }
                else
                    current.Append(c);
            }
            else
            {
                if (c == '"')
                    inQuotes = true;
                else if (c == ',')
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                else
                    current.Append(c);
            }
        }
        result.Add(current.ToString());
        return result.ToArray();
    }
}
