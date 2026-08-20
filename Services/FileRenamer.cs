using ExcelBatchRenamer.Models;
using ExcelBatchRenamer.Utils;

namespace ExcelBatchRenamer.Services;

/// <summary>
/// 文件改名核心引擎
/// </summary>
public class FileRenamer
{
    /// <summary>
    /// 预览改名（只检查文件是否存在，不实际操作）
    /// </summary>
    public List<RenameResult> Preview(RenameConfig config)
    {
        var results = new List<RenameResult>();
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var item in config.Items)
        {
            string oldName = EnsureExtension(item.OldFileName, config.Extension);
            string srcPath = Path.Combine(config.SourceFolder, oldName);

            var result = new RenameResult
            {
                RowNumber = item.RowNumber,
                OldName = oldName,
                OldPath = srcPath,
                IsPreview = true
            };

            if (!File.Exists(srcPath))
            {
                result.Success = false;
                result.Message = "文件缺失";
                results.Add(result);
                continue;
            }

            string newName = BuildNewName(item, config, usedNames);
            if (string.IsNullOrEmpty(newName))
            {
                result.Success = false;
                result.Message = "重名已跳过";
                results.Add(result);
                continue;
            }

            usedNames.Add(newName);
            result.NewName = newName;
            result.NewPath = Path.Combine(config.TargetFolder, newName);
            result.Success = true;
            result.Message = "OK";
            results.Add(result);
        }
        return results;
    }

    /// <summary>
    /// 执行正式改名/复制
    /// </summary>
    public List<RenameResult> Execute(RenameConfig config, IProgress<int>? progress = null)
    {
        var results = new List<RenameResult>();
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int total = config.Items.Count;
        int done = 0;

        // 复制模式下确保输出目录存在
        if (config.Mode == OperationMode.Copy && !string.IsNullOrEmpty(config.OutputFolder))
            Directory.CreateDirectory(config.OutputFolder);

        foreach (var item in config.Items)
        {
            string oldName = EnsureExtension(item.OldFileName, config.Extension);
            string srcPath = Path.Combine(config.SourceFolder, oldName);

            var result = new RenameResult
            {
                RowNumber = item.RowNumber,
                OldName = oldName,
                OldPath = srcPath
            };

            if (!File.Exists(srcPath))
            {
                result.Success = false;
                result.Message = "文件缺失";
                results.Add(result);
                done++;
                progress?.Report(done * 100 / Math.Max(1, total));
                continue;
            }

            string newName = BuildNewName(item, config, usedNames);
            if (string.IsNullOrEmpty(newName))
            {
                result.Success = false;
                result.Message = "重名已跳过";
                results.Add(result);
                done++;
                progress?.Report(done * 100 / Math.Max(1, total));
                continue;
            }

            usedNames.Add(newName);
            string destPath = Path.Combine(config.TargetFolder, newName);

            try
            {
                if (config.Mode == OperationMode.Copy)
                    File.Copy(srcPath, destPath, overwrite: false);
                else
                    File.Move(srcPath, destPath);

                result.NewName = newName;
                result.NewPath = destPath;
                result.Success = true;
                result.Message = "OK";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"错误: {ex.Message}";
            }

            results.Add(result);
            done++;
            progress?.Report(done * 100 / Math.Max(1, total));
        }
        return results;
    }

    /// <summary>
    /// 构建新文件名，处理补零、连接符、清洗、重名
    /// </summary>
    private string BuildNewName(RenameItem item, RenameConfig config, HashSet<string> usedNames)
    {
        // 前缀补零
        string prefix = item.Prefix;
        if (config.PadDigits > 0 && int.TryParse(prefix, out int num))
            prefix = num.ToString($"D{config.PadDigits}");

        // 旧名去掉扩展名
        string baseName = Path.GetFileNameWithoutExtension(item.OldFileName);

        // 组合新文件名
        string newName = $"{prefix}{config.Separator}{baseName}.{config.Extension.TrimStart('.')}";

        // 自动清洗
        if (config.AutoClean)
            newName = FileNameCleaner.Clean(newName);

        // 重名处理
        if (usedNames.Contains(newName))
        {
            if (config.DupMode == DuplicateMode.Skip)
                return string.Empty;

            // 自动加序号
            string nameWithoutExt = Path.GetFileNameWithoutExtension(newName);
            string ext = Path.GetExtension(newName);
            int counter = 1;
            string candidate;
            do
            {
                candidate = $"{nameWithoutExt}({counter}){ext}";
                counter++;
            } while (usedNames.Contains(candidate));
            newName = candidate;
        }

        return newName;
    }

    /// <summary>
    /// 确保文件名带有指定扩展名
    /// </summary>
    private string EnsureExtension(string fileName, string extension)
    {
        if (string.IsNullOrEmpty(Path.GetExtension(fileName)))
            return $"{fileName}.{extension.TrimStart('.')}";
        return fileName;
    }
}
