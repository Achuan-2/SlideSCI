# SlideSCI 签名证书

`SlideSCI-CodeSigning-2026.cer` 是公开证书，不包含私钥。发布者为 `CN=Achuan-2 SlideSCI`，有效期为 2026-10-08 至 2036-10-09，SHA-1 指纹为 `E3DBC9C6ED5D5928755F1C2B8A13FE05215DA6E1`。

VSTO 应用清单和部署清单使用这个证书签名，并通过 DigiCert 时间戳服务记录签名时间。项目根据指纹从 Windows 当前用户的“个人”证书库中读取私钥。更换开发电脑时，需要从原电脑导出带密码的 PFX，并在新电脑的当前用户“个人”证书库中导入。PFX 和密码不得提交到仓库或放入安装包。

Advanced Installer 工程从 `SlideSCI/bin/Debug` 取文件。重新构建项目后，更新安装包版本并重新生成 EXE；EXE 也需要独立签名，构建 VSTO 项目不会自动给安装 EXE 签名。可在 Windows SDK 的 `signtool.exe` 所在环境中执行以下命令，将最后一个参数替换为新安装包路径：

```powershell
signtool.exe sign /sha1 E3DBC9C6ED5D5928755F1C2B8A13FE05215DA6E1 /s My /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 /d SlideSCI "SlideSCI_v2.2.2.exe"
```

程序集强名称仍使用原有的 `SlideSCI_TemporaryKey.pfx`，保持程序集公钥身份。强名称与 VSTO 清单的发布者证书用途不同，原证书过期不影响该密钥用于程序集强名称签名。

自签名证书不会自动成为其他电脑上的受信任发布者。如果用户电脑因发布者不受信任而拒绝加载，在确认下载来源和上述指纹后，可以将这个公开证书导入当前用户的“受信任的根证书颁发机构”和“受信任的发布者”。这一步与证书过期是不同的问题，不需要关闭 Office 的安全检查。
