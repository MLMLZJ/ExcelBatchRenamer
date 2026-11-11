# 代码签名指南

## 一、先说结论：签名到底能解决什么

VirusTotal 上剩下两家给新版报警：

| 引擎 | 判定 | 性质 |
| --- | --- | --- |
| SecureAge | Malicious | 按"未签名可执行文件"整类判 |
| Trapmine | Malicious.moderate.ml.score | 纯机器学习打分 |

> ⚠️ **重要更正（2026-11-11）**：本节原本写"这两家只看签名、加了签名大概率变绿"，
> 那是我的推断，**没有证据支撑，而且现有证据不支持**。
> 详见 [关于签名假设的更正.md](关于签名假设的更正.md)。
> 请先做完那里的对照实验，再决定要不要花钱买证书。

**已知的确定事实**：

| 引擎 | 旧版(34MB PyInstaller) | 新版(61KB C#) |
| --- | --- | --- |
| Trapmine | **Undetected** | **Malicious.moderate.ml.score** |
| SecureAge | Malicious | Malicious |

- 旧版也是未签名的，却通过了 Trapmine → **Trapmine 不是"未签名就判毒"**
- 新版比旧版小 560 倍反而被 Trapmine 判得更可疑 → 判定跟内容有关
- 所以"签名能解决这两家"目前**只是一个待验证的假说**

**目前不能确定**：这两家具体因为什么报警；签名之后会不会变绿。

**正确的做法**：先用 `发布/对照实验/Hello.exe`（空白窗体、未签名）测一下——
如果连它都报，签名才可能有用；如果它过关，那问题在内容，签名救不了。

---

## 二、签名的三条路，按你的情况选

### 路线 A：内网自己造证书（免费，5 分钟，推荐先用这个）

适合：**只在公司内部/自己电脑上用**。

原理是自己当"证书颁发机构(CA)"，把根证书装到目标电脑的"受信任的根证书颁发机构"里，
之后所有用这个 CA 签的程序都被信任。

**缺点**：VirusTotal 上的 SecureAge / Trapmine 仍然会报，因为它们在互联网上，
不认识你的内网 CA。**但对内网杀软（如果有内网管控台）加白就完全够用。**

### 路线 B：买 OV 代码签名证书（几百到两千元/年）

适合：**要发给外部、要过 Windows SmartScreen、要提升整体信誉**。

买来的证书由 DigiCert / Sectigo / GlobalSign 这类公共 CA 签发，
Windows SmartScreen 和各安全厂商都认。

**但请注意**：它能否让 SecureAge / Trapmine 变绿，**目前未经证实**
（见文件开头的更正）。它对 SmartScreen 的效果是明确的，
对这两家 VT 引擎的效果是推测。**建议先做对照实验再决定是否购买。**

如果能确认签名有效，那这确实是唯一能让那两家闭嘴的办法——
但要先验证，不要先花钱。

### 路线 C：内网 CA + 分发根证书（企业环境最标准）

适合：**公司有 IT 部门、有一批电脑要部署**。

让 IT 用公司域控的 CA 签，然后通过组策略把根证书下发到所有电脑。
效果等同路线 B，但不用花钱。**需要 IT 配合。**

> 建议顺序：先用路线 A 验证流程跑通（不花钱），确认没问题后再决定要不要买证书。

---

## 三、路线 A 实操：内网自签名（现在就能做）

### 步骤 1：创建代码签名证书

以**管理员身份**打开 PowerShell，执行：

```powershell
# 创建一张 5 年有效的代码签名证书
$cert = New-SelfSignedCertificate `
    -Type CodeSigningCert `
    -Subject "CN=你的公司名, O=你的公司名, C=CN" `
    -CertStoreLocation "Cert:\LocalMachine\My" `
    -NotAfter (Get-Date).AddYears(5) `
    -KeyUsage DigitalSignature `
    -KeyExportPolicy Exportable

# 记下指纹
$cert.Thumbprint
```

### 步骤 2：导出成 PFX 文件（备份 + 换电脑用）

```powershell
$pwd = ConvertTo-SecureString -String "自己设一个强密码" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath "D:\codesign.pfx" -Password $pwd
```

⚠️ **PFX 文件 + 密码 = 你的身份**，丢了别人就能冒充你签名。
   存到安全的地方，别放共享盘。

### 步骤 3：把根证书装到目标电脑（关键一步）

导出证书（不含私钥）：

```powershell
Export-Certificate -Cert $cert -FilePath "D:\codesign.cer"
```

然后**在每一台要运行这个 exe 的电脑上**，管理员 PowerShell 执行：

```powershell
Import-Certificate -FilePath "D:\codesign.cer" `
    -CertStoreLocation "Cert:\LocalMachine\Root"
```

企业环境用组策略下发更省事：
计算机配置 → Windows 设置 → 安全设置 → 公钥策略 → 受信任的根证书颁发机构 → 导入

### 步骤 4：签名

**方式一：双击脚本（我写好了）**

把 `ExcelPdfRename.exe` 和 `sign.bat` 放在同一个文件夹，双击 `sign.bat`，
按提示选 1，输入 PFX 路径和密码。

**方式二：命令行**

```
"C:\Program Files (x86)\Windows Kits\10\Tools\bin\i386\signtool.exe" ^
  sign /f "D:\codesign.pfx" /p "你的密码" ^
  /fd sha256 /td sha256 /tr http://timestamp.digicert.com ^
  ExcelPdfRename.exe
```

**参数含义（每个都不能少）**：
| 参数 | 作用 |
| --- | --- |
| `/f` `/p` | PFX 文件路径和密码 |
| `/fd sha256` | 文件摘要用 SHA256（不能用 SHA1，已不安全） |
| `/td sha256` | 时间戳摘要也用 SHA256 |
| `/tr` | **RFC3161 时间戳服务器**，见下方说明 |

### 步骤 5：验证

```powershell
Get-AuthenticodeSignature .\ExcelPdfRename.exe | Format-List Status, StatusMessage, SignerCertificate, TimeStamperCertificate
```

看到 `Status : Valid` 就成了。

> 如果显示 `UnknownError` + "证书链在不被信任的根处终止"：
> 说明这台电脑还没装你的根证书（回到步骤 3）。
> 签名本身是好的，装完根证书再看就是 Valid。

---

## 四、时间戳为什么必须有（这一步最容易被漏掉）

不签时间戳：证书 2029 年过期，2029 年之后**所有签过的程序全部变成"签名无效"**，
你需要重新签一遍所有文件。

签了时间戳：时间戳机构证明"这个文件是在 2026 年签的，当时证书有效"。
证书过期后，签名**永久有效**。

所以 `/tr http://timestamp.digicert.com` 不是可选项，是必选项。

可用时间戳服务器（一个不通换另一个）：
- `http://timestamp.digicert.com`
- `http://timestamp.sectigo.com`
- `http://timestamp.globalsign.com/tsa/r6advanced1`

---

## 五、路线 B 实操：买证书

### 买什么

| 类型 | 价格 | 区别 |
| --- | --- | --- |
| **OV**（Organization Validation） | 约 1000~2000 元/年 | 验证公司真实性，够用，**推荐** |
| **EV**（Extended Validation） | 约 3000~5000 元/年 | 更严格，能立刻过 SmartScreen，但个人很难申请 |

个人开发者可以买 OV，但是**2023 年 6 月起，所有代码签名证书的私钥必须存在硬件里**
（USB Key 或云 HSM），不能像以前那样直接给你一个 pfx 文件。

### 这意味着什么

- 买 USB Key 版本：插上 Key，用厂商提供的工具签名，`signtool` 配合 `/n` 或 `/sha1` 挑选证书
- 买云 HSM 版本（DigiCert KeyLocker、Sectigo 等）：用厂商的签名服务，不是本地 signtool

所以**买之前一定要问清楚厂商"怎么集成到 signtool"**，各家工具链不一样。

### 购买后签名（USB Key 版）

```
signtool sign /n "你的公司名" /fd sha256 /td sha256 /tr http://timestamp.digicert.com ExcelPdfRename.exe
```

---

## 六、顺序千万别搞反

**正确顺序：**

```
1. 编译  →  build.bat
2. 签名  →  sign.bat
3. 算哈希 →  certutil -hashfile ExcelPdfRename.exe SHA256
4. 上传 VirusTotal / 提交白名单
```

**为什么**：签名会修改 exe 文件内容（追加证书数据），哈希就变了。
如果先传 VirusTotal 再签名，你上传的哈希和最终用的文件对不上，等于白传。

实测：签名前 `fac7b61b...3baa`（61440 字节），签名后变成别的哈希（68960 字节）。
体积增加了约 7.5 KB，就是证书和时间戳的数据。

---

## 七、签名对"改代码"有没有影响

**没有。** 签名是 exe 编译完成后的独立一步，跟源码无关。

但要注意：**每次重新编译，都要重新签名**（因为编译出来的新文件没签名）。
所以推荐把流程固化成：

```
build.bat   →   编译
sign.bat    →   签名
```

两个都跑完才算完成。我在 `build.bat` 末尾留了提示，跑完会提醒你签名。

---


---

## 九、测试完之后一定要清理（重要）

做对照实验、或测试签名流程时，你会把测试用的根证书装进本机信任库。
**测完必须删掉**，否则等于给系统留了一个"谁拿到这张证书的私钥，
谁签的程序你这台机器都信任"的后门。

### 删测试根证书

```powershell
# 先查有没有（按你的证书名搜）
Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Subject -like "*你的测试证书名*" }

# 确认无误后删除
Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Subject -like "*你的测试证书名*" } | Remove-Item -Force

# 证书本身（含私钥）也一起删
Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -like "*你的测试证书名*" } | Remove-Item -Force

# 复查：应该什么都不输出
Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Subject -like "*你的测试证书名*" }
```

### 要检查的全部位置

| 证书库 | 说明 |
| --- | --- |
| `Cert:\CurrentUser\Root` | **最容易忘、危害最大**，删掉 |
| `Cert:\CurrentUser\My` | 证书本体（带私钥），删掉 |
| `Cert:\LocalMachine\Root` | 如果当初用管理员装到了机器级 |
| `Cert:\CurrentUser\TrustedPublisher` | 有些流程会往这装 |

### 还要删的临时文件

- 测试用的 `.pfx`（带私钥！）
- 测试用的 `.cer`
- 签名过的测试 exe（`Signed.exe` 之类）

> ⚠️ `.pfx` 文件等于你的签名身份，泄露了别人就能冒充你签任何程序。
> 测试用的 PFX 用完立刻删，正式的要存到安全的地方。

### 这条也适用于正式部署

如果你走的是"内网自建 CA"路线，正式用的根证书是要长期装在目标机器上的，
那个**不能删**——但测试用的那张**必须删**，两者别搞混。

---

## 八、常见问题

**Q：签名后杀软还报怎么办？**
A：那说明问题不在签名。把新哈希提交误报申诉（见 `误报申诉模板.txt`）。
   注意：申诉要用**签名后的哈希**。

**Q：一台电脑签了，别的电脑能用吗？**
A：能。签名是写在 exe 文件里的，exe 拷到哪签名就在。但**目标电脑必须信任你的 CA**
   （路线 A 需要装根证书；路线 B 是全球公共 CA，不用装）。

**Q：PFX 文件泄露了怎么办？**
A：立刻吊销证书（联系 CA），然后重新签发一张。路线 A 的话，把根证书从所有电脑删掉
   再重新生成一套。

**Q：能不能不买证书，让 VirusTotal 也全绿？**
A：**不确定**。之前我写"这两家只认公共 CA 签名"，那是未经验证的推断。
   已知 Trapmine 对旧版（未签名）放行、对新版（未签名）报警，
   说明它看的不是签名状态。所以买了证书也未必能变绿。
   建议先做对照实验（见文件开头），再决定是否购买。
   就目前而言，68/70 + 内网加白是成本最低、效果最确定的方案。

**Q：这台机器现在能签吗？**
A：能。我已经确认这台机器有 `signtool.exe`（在 Windows ADK 里），
   而且刚才用一张测试证书完整跑通了签名 + 时间戳流程，
   签名后 `TimeStamperCertificate` 正确显示 DigiCert 时间戳。
   你现在只差一张自己的证书（路线 A 的步骤 1 就能生成）。