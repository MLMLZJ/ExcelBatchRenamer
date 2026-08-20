namespace ExcelBatchRenamer.Models;

public enum OperationMode
{
    /// <summary>直接改名（原文件被修改）</summary>
    Rename,
    /// <summary>复制到新文件夹（原文件保留）</summary>
    Copy
}

public enum DuplicateMode
{
    /// <summary>跳过重名文件</summary>
    Skip,
    /// <summary>自动加序号 (1)(2)</summary>
    AutoNumber
}

/// <summary>
/// 改名操作配置
/// </summary>
public class RenameConfig
{
    /// <summary>源文件所在文件夹</summary>
    public string SourceFolder { get; set; } = string.Empty;

    /// <summary>输出文件夹（复制模式时使用）</summary>
    public string OutputFolder { get; set; } = string.Empty;

    /// <summary>改名映射项列表</summary>
    public List<RenameItem> Items { get; set; } = new();

    /// <summary>前缀和旧名之间的连接符，支持中文</summary>
    public string Separator { get; set; } = "_";

    /// <summary>目标扩展名（不带点，如 pdf）</summary>
    public string Extension { get; set; } = "pdf";

    /// <summary>前缀补零位数，0=不补零</summary>
    public int PadDigits { get; set; } = 0;

    /// <summary>重名处理方式</summary>
    public DuplicateMode DupMode { get; set; } = DuplicateMode.Skip;

    /// <summary>操作模式</summary>
    public OperationMode Mode { get; set; } = OperationMode.Rename;

    /// <summary>是否自动清洗文件名（去空格、替换非法字符）</summary>
    public bool AutoClean { get; set; } = true;

    /// <summary>目标文件夹（用于写撤销记录）</summary>
    public string TargetFolder => Mode == OperationMode.Copy ? OutputFolder : SourceFolder;
}
