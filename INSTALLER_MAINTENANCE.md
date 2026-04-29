# SlideSCI Installer Maintenance

本文件记录 `SlideSCI_WPS_PowerPoint_Compat` 安装包的已验证成功流程，避免后续升级安装器时再次引入“重复安装后插件不显示”或“安装器双击无界面”的问题。

## 当前成功基线

- 成功版本：`1.4.6.24`
- 安装包：`artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.24.exe`
- 验证结果：
  - 双击安装器能正常弹出界面。
  - 重复安装同一版本后 PowerPoint 和 WPS 插件界面仍正常显示。
  - PowerPoint Addins `LoadBehavior` 保持 `3`。
  - 安装后 `{app}` 根目录存在 `SlideSCICompat.dll`、`DocumentFormat.OpenXml.dll`、`Newtonsoft.Json.dll` 等运行时文件。

## 不要再做的事

- 不要在 `InitializeSetup` 中运行 `tasklist` 或其他外部进程检测。该阶段发生在安装器窗口创建前，曾导致安装器双击后无界面，后台堆积多个 `.tmp` 进程。
- 不要在安装流程中 `ShellExec` 打开 `{app}\SlideSCICompat.vsto` 并等待退出。该流程曾导致 Inno Setup 卡住。
- 不要手写当前 `HKCU\Software\Microsoft\VSTO\SolutionMetadata\{SolutionId}` 的完整内容。该键应尽量交给 VSTO Runtime/VSTOInstaller 维护。
- 不要只依赖 `Application Files\SlideSCICompat_x_y_z` 子目录中的 DLL。当前 VSTO 加载路径需要 `{app}` 根目录也能解析主程序集和依赖。

## 必须保留的安装流程

1. 安装界面先正常显示。
2. 在 `PrepareToInstall` 阶段关闭宿主程序：
   - `POWERPNT`
   - `WPP`
   - `WPS`
3. 文件复制完成后执行 `ExpandDeployFiles`，将 `.deploy` 展开为真实文件。
4. 执行 `CopyRuntimeFilesToAppRoot`，把当前版本目录中的 `.dll`、`.config`、`.manifest` 复制到 `{app}` 根目录。
5. 执行 `RegisterVstoTrust`，写入：
   - `HKCU\Software\Microsoft\Office\PowerPoint\Addins\SlideSCICompat`
   - `HKCU\Software\Kingsoft\Office\WPP\AddinsWL`
   - `HKCU\Software\Microsoft\VSTO\Security\Inclusion`
   - `LoadBehavior=3`
6. 安装结束阶段再次执行 `RegisterVstoTrust`，确保最终注册状态收口。

## 每次升级版本必须修改

- `SlideSCI/SlideSCI.csproj`
  - `<ApplicationVersion>x.y.z.n</ApplicationVersion>`
- `SlideSCI/Properties/AssemblyInfo.cs`
  - `AssemblyVersion`
  - `AssemblyFileVersion`
- `CODER_LOG.md`
  - 记录代码改动、构建产物、验证结果

## 构建命令

```powershell
& 'C:\Program Files\PowerShell\7\pwsh.exe' -NoLogo -NoProfile -ExecutionPolicy Bypass -File 'D:\SlideSCI_wps\tools\build-release.ps1'
```

## 安装验证命令

使用静默安装验证安装器不会卡住：

```powershell
$log = 'D:\SlideSCI_wps\artifacts\dist\install-smoke.log'
Remove-Item $log -Force -ErrorAction SilentlyContinue
& 'D:\SlideSCI_wps\artifacts\dist\SlideSCI_WPS_PowerPoint_Compat_v1.4.6.24.exe' /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG=$log
Get-Content $log -Tail 40
```

日志中必须出现：

```text
Installation process succeeded.
```

检查注册表：

```powershell
reg query HKCU\Software\Microsoft\Office\PowerPoint\Addins\SlideSCICompat /s
reg query HKCU\Software\Kingsoft\Office\WPP\AddinsWL /v SlideSCICompat
reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\{0E5CB4DB-3A3D-4AB8-B8CC-561DD944451A}_is1" /v DisplayName
```

期望：

```text
LoadBehavior    REG_DWORD    0x3
DisplayName     REG_SZ       SlideSCI WPS PowerPoint version <当前版本>
```

检查根目录运行时文件：

```powershell
Get-ChildItem 'D:\SlideSCI_WPS_PowerPoint_Compat' -File |
  Where-Object { $_.Name -in 'SlideSCICompat.dll','DocumentFormat.OpenXml.dll','Newtonsoft.Json.dll' } |
  Select-Object Name,Length,LastWriteTime
```

## 重复安装验证

每次发布前至少执行：

1. 双击安装包，确认安装界面能弹出。
2. 完成安装后打开 PowerPoint，确认插件界面显示。
3. 完成安装后打开 WPS，确认插件界面显示。
4. 关闭 PowerPoint/WPS。
5. 重复安装同一个安装包 2 次。
6. 再次打开 PowerPoint/WPS，确认插件界面仍显示。
7. 检查没有残留安装器进程：

```powershell
Get-Process | Where-Object { $_.ProcessName -like 'SlideSCI_WPS_PowerPoint_Compat_v*' }
```

## 已定位过的故障和根因

- 重复安装后插件消失：
  - 表现：`LoadBehavior` 变为 `2`。
  - 处理：安装末尾再次执行 `RegisterVstoTrust`，确保最终 `LoadBehavior=3`。
- Office 报 `未能加载文件或程序集 "SlideSCICompat"`：
  - 表现：VSTO 找不到主程序集。
  - 根因：VSTO 加载阶段只从 `{app}` 根目录探测主程序集。
  - 处理：保留 `CopyRuntimeFilesToAppRoot`，把当前版本目录的运行时 DLL 同步到安装根目录。
- 安装器双击无界面：
  - 表现：后台出现多个安装器 `.tmp` 进程，没有 UI。
  - 根因：`InitializeSetup` 中运行 `tasklist` 卡在窗口创建前。
  - 处理：不要在 `InitializeSetup` 中执行外部进程检测；改用 `PrepareToInstall`。
