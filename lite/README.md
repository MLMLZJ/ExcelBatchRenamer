# Lite 版（ExcelPdfRename.exe）

这是同一工具的另一套实现，和仓库根目录的 **Pro 版**（ExcelBatchRenamer）并存，
两者功能一致但技术栈和适用场景不同。

## 为什么有两个版本

Pro 版是 `.NET 8 + NPOI`，Lite 版是 `.NET Framework 4.8 + 原生 xlsx 解析`。
不是为了炫技，是因为一个具体问题：**杀软误报**。

最初的 Python + PyInstaller 版本被 VirusTotal 上 3 家引擎报毒，
排查后发现命中的是 PyInstaller 的自解压引导器。于是有了这套重写。
完整实验记录见 [../docs/antivirus-experiments.md](../docs/antivirus-experiments.md)。

## 两个版本怎么选

| | Pro（仓库根目录） | Lite（本目录） |
| --- | --- | --- |
| 运行时 | .NET 8（自包含版 76 MB / 轻量版 17.5 MB） | **.NET Framework 4.8（Windows 自带）** |
| 依赖 | NPOI 2.7.1 | **无第三方依赖** |
| 产物体积 | 17.5 ~ 76 MB | **69 KB** |
| 启动速度 | 较快 | **毫秒级** |
| 目标机器要求 | 需 .NET 8 运行时（或自包含版） | **Windows 10/11 开箱即用** |
| 代码签名 | 需自行签名 | **已签名，VirusTotal 全通过** |
| 源码结构 | 分层（Models / Services / Utils） | 单文件引擎（Core.cs） |

**建议**：

- 要发给别人、对方电脑环境不确定 → **Lite**（零依赖，双击就跑）
- 自己要二次开发、想改架构 → **Pro**（分层清晰，.NET 8 现代语法）
- 需要过杀软 → **Lite**（已实测通过，Pro 未做签名验证）

## 编译

需要 .NET Framework 4.x（Windows 10/11 已内置）。

```bat
build\build.bat
```

产物：`dist\ExcelPdfRename.exe`

签名：

```bat
build\sign.bat
```

## 测试

```bat
tests\run-tests.bat
```

73 项断言，三套：

| 测试 | 断言数 | 覆盖 |
| --- | --- | --- |
| `TestRunner.cs` | 32 | Excel 读取、命名规则、预览/执行/撤销、批量改名 |
| `EdgeTest.cs` | 20 | 空行、空单元格、非法行范围、列名不存在、扩展名边界 |
| `FixTest.cs` | 21 | 大小写改名、只读文件复制、rPh 注音、重名表头、正则反向引用 |

前两套需要先生成测试数据（Python + openpyxl）：

```bat
python tests\make_testdata.py
python tests\make_edge.py
```

`FixTest` 自带临时数据，不依赖外部文件。

## 目录

```
lite/
├── src/
│   ├── Core.cs           # Excel 读取（原生 zip+XML）、命名规则、预览/执行/撤销引擎
│   ├── UiKit.cs          # 控件构造工具（统一排版、高 DPI 适配）
│   ├── MainForm.cs       # 主界面
│   ├── BatchForm.cs      # 普通批量改名窗口
│   ├── Program.cs        # 入口 + 全局异常处理
│   └── AssemblyInfo.cs
├── resources/
│   ├── app.ico
│   └── app.manifest
├── build/
│   ├── build.bat
│   └── sign.bat
└── tests/
    ├── TestRunner.cs
    ├── EdgeTest.cs
    ├── FixTest.cs
    ├── run-tests.bat
    └── make_*.py
```

## 实现上的几个要点

**不用 NPOI**：直接解析 xlsx（zip + XML）。冷启动从数秒降到毫秒级，
产物从 MB 级降到 69 KB。代价是要自己处理 OOXML 细节。

**rPh 注音**：读取单元格时必须跳过 `<rPh>` 节点（日文/中文注音、拼音），
否则开了「显示拼音」的表格会读出脏字符串，导致文件名匹配失败。

**大小写改名**：Windows 文件名大小写不敏感，`report.pdf` → `REPORT.PDF`
会被 `File.Exists` 误判为「目标已存在」而跳过。代码里对「仅大小写不同」单独放行。

**只读源文件**：复制时如果源文件只读，副本也是只读，设置时间戳会抛异常，
造成「复制成功却报失败、还漏写记录」。现在先清只读位再设时间戳。

**表头重名**：Excel 里两列都叫「编号」时，下拉框会分不清。
加载时自动加 `(2)` `(3)` 后缀。

## 已知边界

- 长路径（超过 260 字符）未做 `\\?\` 前缀处理
- 日期单元格读出的是 Excel 序列号（如 45292）
- 错误值单元格（`#DIV/0!`）会原样当文件名
- 表头恒取 Excel 第 1 行

## License

MIT
