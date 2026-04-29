# Coder Log

用于记录当前协作过程中的代码与构建改动，不替代产品发布用的 `CHANGELOG.md`。

## 2026-04-29

### 1.4.6.24

- 升级版本号到 `1.4.6.24`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复安装包双击后无界面的问题
  - 根因：`1.4.6.16` 起引入的 `InitializeSetup -> tasklist` 进程检测会卡在安装器初始化阶段，窗口尚未创建，因此双击表现为无界面
  - 证据：卡住的临时目录只有 `_isetup` 和空的 `slidesci_process_check.txt`，并且 `1.4.6.14` 无该检测逻辑所以能弹窗
  - 移除 `InitializeSetup` 中基于 `tasklist` 的 Office/WPS 进程检测
  - 改为 `PrepareToInstall` 阶段调用 PowerShell `Stop-Process -Name POWERPNT,WPP,WPS`，确保安装界面先显示，再在正式安装前自动关闭宿主程序
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.24.exe`
- 安装验证
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.24`，日志显示 `Installation process succeeded`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.24`
  - 安装后无残留 `SlideSCI_WPS_PowerPoint_Compat_v1.4.6.*` 安装器进程

### 1.4.6.23

- 升级版本号到 `1.4.6.23`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复 `1.4.6.22` 双击安装包无界面/后台卡住的问题
  - 现象：系统中出现多个 `SlideSCI_WPS_PowerPoint_Compat_v1.4.6.22` 和 `.tmp` 安装进程
  - 原因：安装器在 `ssPostInstall` 阶段 `ShellExec` 打开 `.vsto` 并等待退出，容易卡住 Inno 安装流程
  - 移除安装阶段打开 `.vsto` 的逻辑
  - 保留 `CopyRuntimeFilesToAppRoot`，通过将主 DLL 和依赖 DLL 放到安装根目录解决 VSTO `Assembly.Load` 找不到主程序集的问题
  - 安装器只负责复制运行时文件、写 PowerPoint/WPS 注册项、写 VSTO trust 和 `LoadBehavior=3`
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.23.exe`

### 1.4.6.22

- 升级版本号到 `1.4.6.22`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 根因修复 PowerPoint/WPS 加载插件失败
  - 验证：手动将 `Application Files\SlideSCICompat_1_4_6_21` 下的主 DLL 和依赖 DLL 复制到安装根目录后，PowerPoint/WPS 插件恢复显示，`LoadBehavior` 保持 `3`
  - `installer/SlideSCI.iss` 新增 `CopyRuntimeFilesToAppRoot`
  - 安装时自动将当前版本目录下的 `.dll`、`.config`、`.manifest` 复制到 `{app}` 根目录
  - 解决 VSTO Runtime 在 `Assembly.Load("SlideSCICompat, Version=...")` 阶段只从安装根目录探测主程序集导致的 `FileNotFoundException`
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.22.exe`

### 1.4.6.21

- 升级版本号到 `1.4.6.21`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复 VSTO 加载主程序集失败的问题
  - 现象：Office 报错 `未能加载文件或程序集 "SlideSCICompat, Version=1.4.6.20"`
  - 尝试将 ClickOnce 发布属性改为 `<Install>true</Install>`；构建后 `.vsto` 仍为 `deployment install="false"`，说明该属性未被当前 VSTO Publish target 采纳
  - `installer/SlideSCI.iss` 新增 `ResetCurrentVstoMetadata`，安装前清掉当前 VSTO metadata，交给 VSTOInstaller 重新创建
  - `RegisterVstoTrust` 不再删除/手写当前 `HKCU\Software\Microsoft\VSTO\SolutionMetadata\{SolutionId}`，避免破坏 VSTO Runtime 的程序集探测路径
  - 安装器仍负责写 PowerPoint Addins、WPS 白名单、VSTO Inclusion trust 和 `LoadBehavior=3`
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.21.exe`

### 1.4.6.20

- 升级版本号到 `1.4.6.20`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复 `1.4.6.19` 仍不能稳定显示插件的问题
  - 验证发现直接调用 `VSTOInstaller.exe /install` 后，PowerPoint Addins `LoadBehavior` 仍保持 `2`
  - 将安装器改为在 `[Code]` 阶段通过 `ShellExec` 打开 `{app}\SlideSCICompat.vsto`，复刻用户手动双击 `.vsto` 的安装路径
  - 保留 `ShellExec` 后再次调用 `RegisterVstoTrust`，确保同版本重复安装后最终 `LoadBehavior=3`
  - 不再使用 `[Run]` 阶段启动 `.vsto`，避免启动太晚导致注册表无法收口
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.20.exe`

### 1.4.6.19

- 升级版本号到 `1.4.6.19`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复 `1.4.6.18` 移除 `.vsto` 自动执行后 PowerPoint/WPS 插件不显示的问题
  - `installer/SlideSCI.iss` 新增 `FindVstoInstaller`
  - `installer/SlideSCI.iss` 新增 `InstallVstoAddin`
  - 安装阶段同步调用 `VSTOInstaller.exe /install "{app}\SlideSCICompat.vsto"`，确保 ClickOnce/VSTO 本机安装缓存被创建
  - `VSTOInstaller` 执行后再次调用 `RegisterVstoTrust`，避免同版本重复安装把 `LoadBehavior` 留在 `2`
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.19.exe`

### 1.4.6.18

- 升级版本号到 `1.4.6.18`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复同版本重复安装后 PowerPoint/WPS 插件界面消失的问题
  - 移除安装完成后自动执行 `{app}\SlideSCICompat.vsto` 的逻辑
  - 原因：同版本重复安装时，外部 VSTOInstaller 可能把 Office Addins `LoadBehavior` 回写为 `2`
  - 现在由安装器内部 `RegisterVstoTrust` 统一写入 VSTO trust、Manifest 和 `LoadBehavior=3`
  - 修正 `installer/SlideSCI.iss` Code 段注释语法，确保 Inno Setup 可编译
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.18.exe`

### 1.4.6.17

- 升级版本号到 `1.4.6.17`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 调整安装器宿主进程处理
  - 安装开始时自动关闭 `POWERPNT.EXE`、`WPP.EXE`、`WPS.EXE`
  - 如果关闭后仍检测到进程运行，则中止安装并提示手动关闭
- 修复重复安装后插件界面消失/旧版本仍加载的收口问题
  - 在安装完成阶段再次调用 `RegisterVstoTrust`
  - 确保 `LoadBehavior` 最终回写为 `3`

### 1.4.6.16

- 升级版本号到 `1.4.6.16`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复安装时宿主程序仍在运行导致旧 VSTO 缓存继续加载的问题
  - `installer/SlideSCI.iss` 新增安装前进程检测
  - 检测 `POWERPNT.EXE`、`WPP.EXE`、`WPS.EXE`
  - 如果 PowerPoint 或 WPS 仍在运行，则中止安装并提示先完全退出

### 1.4.6.15

- 升级版本号到 `1.4.6.15`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复安装器重复安装后 WPS 插件界面可能消失的问题
  - 在 `installer/SlideSCI.iss` 中新增 `CleanupWpsAddinState`
  - 清理 `WPP\\AddinsWL`、`WPP\\AddinsBL`、`WPP\\AddinsCL`、`6.0\\WPP\\AddIns` 中的 `SlideSCI` / `SlideSCICompat` 残留
  - 注册前增加 `CleanupStalePowerPointAddinPaths` 与 `CleanupWpsAddinState`
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.15.exe`

### 1.4.6.14

- 升级版本号到 `1.4.6.14`
- 调整 `LatexToSvgConverter`
  - 不再在启动时固定缓存脚本路径
  - 每次转换时重新解析可用的 `latex-to-svg.js` 路径
  - 优先选择带完整 `node_modules/mathjax-full` 的目录
- 调整 `latex-converter/latex-to-svg.js`
  - 回退为更接近 MathJax 原生输出的 SVG 生成方式
  - 保留 `xrightarrow/xleftarrow` 等常见输入归一化
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.14.exe`

### 1.4.6.13

- 升级版本号到 `1.4.6.13`
- 修复运行时脚本路径不生效问题
  - `SlideSCI/LatexToSvgConverter.cs` 从硬编码 `D:\\SlideSCI_WPS_PowerPoint_Compat\\latex-converter` 改为多路径候选解析
  - 增加 `%APPDATA%\\Achuan-2\\SlideSCI\\latex-converter`、安装目录、程序集目录等候选路径
  - 在错误信息中输出已检查路径
- 增加开发机 `latex-converter` 自动同步
  - `SlideSCI/SlideSCI.csproj` 在 `Build` 后自动同步 `latex-to-svg.js`、`package.json`、`package-lock.json` 到 `D:\\SlideSCI_WPS_PowerPoint_Compat\\latex-converter`
  - 将 `package-lock.json` 纳入发布内容
- 改进 LaTeX 输入预处理
  - `SlideSCI/Ribbon1.cs` 统一 LaTeX 输入归一化
  - 自动识别多行公式并包装为 `aligned`
  - 对 `xrightarrow/xleftarrow` 等常见输入进行补全
  - 对包含 `overline/underbrace/xrightarrow/...` 的公式优先走 SVG
- 改进 `latex-converter/latex-to-svg.js`
  - 增加常见命令纠错
  - 增加 SVG 后处理与嵌套节点兼容逻辑
- 修复本机注册
  - 执行 `tools/register_addin.ps1`，补齐 `|vstolocal` 等注册项
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.13.exe`
