using System.Text.RegularExpressions;

namespace ExcelBatchRenamer.Utils;

/// <summary>
/// 文件名清洗工具
/// </summary>
public static class FileNameCleaner
{
    /// <summary>Windows文件名非法字符</summary>
    private static readonly char[] InvalidChars = { '\\', '/', ':', '*', '?', '"', '<', '>', '|' };

    /// <summary>
    /// 清洗文件名：去除前后空格、替换非法字符为下划线
    /// </summary>
    public static string Clean(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            return fileName;

        // 去除前后空格
        fileName = fileName.Trim();

        // 替换非法字符
        foreach (var c in InvalidChars)
            fileName = fileName.Replace(c, '_');

        // 合并连续空格
        fileName = Regex.Replace(fileName, @"\s+", " ");

        return fileName;
    }

    /// <summary>
    /// 检查文件名是否包含非法字符
    /// </summary>
    public static bool HasInvalidChars(string fileName)
    {
        return fileName.IndexOfAny(InvalidChars) >= 0;
    }
}
