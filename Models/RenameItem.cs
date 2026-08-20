namespace ExcelBatchRenamer.Models;

/// <summary>
/// 单条改名映射项
/// </summary>
public class RenameItem
{
    /// <summary>Excel行号（含表头，从1开始）</summary>
    public int RowNumber { get; set; }

    /// <summary>前缀（如 FL01、001）</summary>
    public string Prefix { get; set; } = string.Empty;

    /// <summary>旧文件名（可带或不带扩展名）</summary>
    public string OldFileName { get; set; } = string.Empty;
}
