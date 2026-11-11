# Excel 批量文件重命名工具

Windows 桌面批量文件重命名工具，读取 Excel 映射表按规则改名。
提供两套实现，功能一致，按使用场景选择。

| | **Pro 版**（仓库根目录） | **Lite 版**（[`lite/`](lite/)） |
| --- | --- | --- |
| 运行时 | .NET 8 | .NET Framework 4.8（Windows 自带） |
| 依赖 | NPOI 2.7.1 | 无第三方依赖 |
| 产物体积 | 17.5 MB（轻量）/ 76 MB（自包含） | **69 KB** |
| 启动速度 | 较快 | **毫秒级** |
| 目标机器要求 | 需 .NET 8 运行时，或发自包含版 | **Windows 10/11 开箱即用** |
| 代码签名 | 需自行签名 | **已签名，VirusTotal 全通过** |
| 源码结构 | 分层（Models / Services / Utils） | 单文件引擎（Core.cs） |

**怎么选**：

- 发给别人、对方环境不确定 → **Lite**（零依赖，双击就跑）
- 自己二次开发、要清晰架构 → **Pro**（分层，现代 C#）
- 需要过杀软 → **Lite**（已实测通过）

---

## 功能

### Excel 映射改名

读取 `.xlsx` / `.xlsm` 第一个工作表，第一行为表头。

- 选择「前缀列」和「旧名字列」，生成 `前缀 + 连接符 + 旧名字.扩展名`
- 连接符支持中文（`_` `-` `第` `号_` …），扩展名自定义
- 前缀补零（`1` → `001`）
- 行范围筛选（行号含表头行）
- 重名处理：跳过 / 自动加序号
- 自动清洗文件名（替换 `\ / : * ? " < > |`）
- 两种模式：直接改名 / 复制到新文件夹

### 普通批量改名

- 文件类型过滤（`*.pdf` `*.docx` `*.*`，多个用空格分隔）
- 加前缀 / 加后缀、查找替换（支持正则）、大小写转换
- 执行顺序：查找替换 → 大小写转换 → 加前缀 → 加后缀

### 安全特性

- **预览模式**：先看结果，不动文件
- **执行前二次确认**，有缺失/重名时额外提示
- **每次执行生成记录 CSV**（UTF-8 BOM，Excel 直接可读）
- **一键撤销**：改名模式改回原名，复制模式删除副本
- 彩色日志、进度条、日志导出

---

## Pro 版

### 技术栈

- **.NET 8** (LTS)
- **WinForms**
- **NPOI 2.7.1** — Excel 读取
- **C# 12**

### 编译发布

自包含版（76 MB，任何 Windows 电脑都能跑）：

```bash
dotnet publish -c Release -r win-x64 --self-contained true ^
  /p:PublishSingleFile=true ^
  /p:IncludeNativeLibrariesForSelfExtract=true ^
  /p:EnableCompressionInSingleFile=true
```

轻量版（17.5 MB，需 .NET 8 运行时）：

```bash
dotnet publish -c Release -r win-x64 --self-contained false /p:PublishSingleFile=true
```

产物在 `bin/Release/net8.0-windows/win-x64/publish/`。

### 项目结构

```
.
├── Program.cs
├── MainForm.cs
├── BatchRenameForm.cs
├── Models/
│   ├── RenameItem.cs
│   ├── RenameConfig.cs
│   └── RenameResult.cs
├── Services/
│   ├── ExcelReader.cs
│   ├── FileRenamer.cs
│   └── UndoManager.cs
└── Utils/
    └── FileNameCleaner.cs
```

---

## Lite 版

完整说明见 [`lite/README.md`](lite/README.md)。

### 为什么会有 Lite 版

最初这个工具是 Python + Tkinter 写的，用 `pyinstaller -F` 打包成 34 MB 单文件 exe。
它在 VirusTotal 上被 3 家引擎报毒：

| 引擎 | 判定 |
| --- | --- |
| Bkav Pro | W32.Malware.F31CED00 |
| Elastic | Malicious (moderate confidence) |
| SecureAge | Malicious |

其余 66 家（Microsoft、Kaspersky、ESET、CrowdStrike 等）都不报。

排查后发现命中的不是业务代码，而是 **PyInstaller 的自解压引导器**——
它运行时把 1798 个文件解压到 `%TEMP%\_MEIxxxx`，这个结构本身就是杀软重点关注对象。

于是用 C# / .NET Framework 重写，并做了对照实验验证。实测结果：

| 版本 | 体积 | 签名 | VirusTotal |
| --- | --- | --- | --- |
| Python / PyInstaller | 34 MB | 无 | 3/69 |
| Python / PyInstaller | 34 MB | **有** | **3/68** |
| Lite（C#） | 61 KB | 无 | 2/70 |
| **Lite（C#）** | **69 KB** | **有** | **通过** |

**结论：只签名救不了 PyInstaller 版（结构不变，该报还报），只换 C# 不签名也还有残留——两者缺一不可。**

完整实验记录：[docs/antivirus-experiments.md](docs/antivirus-experiments.md)

### 编译

```bat
cd lite
build\build.bat
```

产物：`lite\dist\ExcelPdfRename.exe`

签名（可选）：

```bat
build\sign.bat
```

详见 [docs/signing-guide.md](docs/signing-guide.md)。

### 测试

```bat
cd lite
tests\run-tests.bat
```

73 项断言。前两套需要先生成测试数据：

```bat
python tests\make_testdata.py
python tests\make_edge.py
```

### 实现要点

**不用 NPOI**：直接解析 xlsx（zip + XML）。冷启动从数秒降到毫秒级，产物从 MB 降到 69 KB。

**rPh 注音**：读取单元格时跳过 `<rPh>` 节点，否则开了「显示拼音」的中日文表格会读出脏字符串。

**大小写改名**：Windows 文件名大小写不敏感，`report.pdf` → `REPORT.PDF` 会被
`File.Exists` 误判为「目标已存在」。对「仅大小写不同」单独放行。

**只读源文件**：复制时源文件只读会导致副本只读，设置时间戳抛异常，
造成「复制成功却报失败、还漏写记录」。现在先清只读位再设时间戳。

---

## 下载

最新 Release：见 [Releases](../../releases) 页面。

- **Lite 版**（推荐直接使用）：`ExcelBatchRename-v3.0-signed-win-x64.zip` — 已签名，69 KB，解压即用
- **Pro 版**：`ExcelBatchRenamer_v2.0.0_x64.exe` / `_lite.exe`

## 系统要求

- Windows 10 / 11 x64
- Pro 版：.NET 8 运行时（自包含版无需）
- Lite 版：无需安装任何东西（.NET Framework 4.8 系统自带）

## 文档

- [代码签名指南](docs/signing-guide.md) — 自签名 / OV 证书 / 时间戳 / 清理
- [杀软误报排查](docs/antivirus-experiments.md) — 6 个 VT 样本的对照实验
- [PyInstaller 签名实测](docs/pyinstaller-signing-result.md) — 为什么签名救不了它
- [中文使用说明](docs/usage-zh.txt)

## License

MIT
