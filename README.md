# Excel批量文件重命名工具 Pro

基于 C# + WinForms + NPOI 开发的 Windows 桌面批量文件重命名工具。

## 功能特性

### Excel 映射改名（主功能）
- 读取 Excel (.xlsx) 映射表，选择前缀列和旧名字列
- 命名规则：`前缀 + 连接符 + 旧名.扩展名`
- 连接符支持中文（如 `_`、`-`、`第`、`号_`）
- 前缀补零（如 1 → 001）
- 行范围筛选（从第 X 行到第 Y 行）
- 重名处理：跳过 / 自动加序号 (1)(2)
- 自动清洗文件名（去空格、替换非法字符）
- 两种操作模式：直接改名 / 复制到新文件夹
- 预览模式（只检查文件是否存在，不实际操作）
- 撤销改名（CSV 记录文件回滚）
- 日志导出

### 普通批量改名（子窗口）
- 文件类型过滤（`*.pdf` `*.docx` 等，多个用空格分隔）
- 添加前缀 / 添加后缀
- 查找替换（支持正则表达式）
- 大小写转换（不变 / 全部大写 / 全部小写 / 首字母大写）
- 改名执行顺序：查找替换 → 大小写转换 → 添加前缀 → 添加后缀
- 支持预览、执行、撤销

## 技术栈

- **.NET 8** (LTS)
- **WinForms** — Windows 原生 GUI
- **NPOI** — Excel (.xlsx) 读取
- **C# 12**

## 系统要求

- Windows 10 / 11 x64
- 自包含版：无需安装任何运行时
- 轻量版：需安装 [.NET 8 运行时](https://dotnet.microsoft.com/download/dotnet/8.0)

## 编译发布

### 自包含版（76MB，任何 Windows 电脑都能跑）
```bash
dotnet publish -c Release -r win-x64 --self-contained true ^
  /p:PublishSingleFile=true ^
  /p:IncludeNativeLibrariesForSelfExtract=true ^
  /p:EnableCompressionInSingleFile=true
```

### 轻量版（17.5MB，需 .NET 8 运行时）
```bash
dotnet publish -c Release -r win-x64 --self-contained false /p:PublishSingleFile=true
```

产物位于 `bin/Release/net8.0-windows/win-x64/publish/`。

## 项目结构

```
ExcelBatchRenamer/
├── Program.cs                   # 程序入口
├── MainForm.cs                  # 主窗口（Excel映射改名）
├── BatchRenameForm.cs           # 普通批量改名子窗口
├── Models/
│   ├── RenameItem.cs            # 改名映射项
│   ├── RenameConfig.cs          # 改名配置
│   └── RenameResult.cs          # 改名结果
├── Services/
│   ├── ExcelReader.cs           # Excel 读取服务
│   ├── FileRenamer.cs           # 文件改名核心引擎
│   └── UndoManager.cs           # 撤销记录管理
└── Utils/
    └── FileNameCleaner.cs       # 文件名清洗工具
```

## 使用说明

1. 双击 `ExcelBatchRenamer.exe` 启动
2. 选择 Excel 映射表（.xlsx，第一行为表头）
3. 选择文件所在文件夹
4. 点击"加载Excel列"
5. 下拉选择前缀列和旧名字列
6. 设置命名规则（连接符、扩展名、补零位数等）
7. 选择操作模式（直接改名 / 复制到新文件夹）
8. 点击"预览改名"确认无误
9. 点击"执行正式改名"

## License

MIT
