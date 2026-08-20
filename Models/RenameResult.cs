namespace ExcelBatchRenamer.Models;

/// <summary>
/// 单条改名结果
/// </summary>
public class RenameResult
{
    /// <summary>Excel行号</summary>
    public int RowNumber { get; set; }

    /// <summary>原文件名</summary>
    public string OldName { get; set; } = string.Empty;

    /// <summary>原文件完整路径</summary>
    public string OldPath { get; set; } = string.Empty;

    /// <summary>新文件名</summary>
    public string NewName { get; set; } = string.Empty;

    /// <summary>新文件完整路径</summary>
    public string NewPath { get; set; } = string.Empty;

    /// <summary>是否成功</summary>
    public bool Success { get; set; }

    /// <summary>状态消息（成功/缺失/错误信息）</summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>是否为预览模式（未实际操作）</summary>
    public bool IsPreview { get; set; }
}
