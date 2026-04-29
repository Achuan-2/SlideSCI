# Coder Log

用于记录当前协作过程中的代码与构建改动，不替代产品发布用的 `CHANGELOG.md`。

## 2026-04-29

### 1.4.6.52

- 升级版本号到 `1.4.6.52`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 项目公开显示名改为 `SlideBridge Office`
  - Ribbon 显示为 `SlideBridge`
  - 安装器显示名改为 `SlideBridge Office`
  - 安装包输出名改为 `SlideBridge_Office_WPS_PowerPoint_v{version}.exe`
- 更新公开个人信息
  - GitHub: `https://github.com/jacywallny/`
  - Email: `jacywalln@gmail.com`
- 清理公开原作者信息
  - 重写 `README.md` 与 `README.en.md`
  - 删除 Ribbon 中的资金支持入口和对应图片资源
  - 替换关于区与 GitHub 按钮链接
- 保留内部兼容标识
  - `SlideSCICompat` assembly/add-in key 不改
  - 旧注册表清理逻辑不改
- 安全与构建检查
  - `latex-converter` 执行 `npm audit --json`，报告 0 个漏洞
  - `tools/build-release.ps1 -Version 1.4.6.51 -SkipInstaller` 执行成功
  - `tools/build-release.ps1 -Version 1.4.6.52` 执行成功
  - 生成安装包 `artifacts/dist/SlideBridge_Office_WPS_PowerPoint_v1.4.6.52.exe`

### 1.4.6.51

- 修复 PowerPoint「插入 LaTeX SVG」处理长正文夹行内公式时被压成超小单行图片的问题
  - PowerPoint 下检测到“正文 + `$...$` 行内公式”混合输入时，改为插入可换行文本框
  - 纯公式输入仍保留 SVG 路径，WPS 路径不变

### 1.4.6.50

- 调整 PowerPoint 插入 LaTeX SVG 的默认自适应尺寸
  - WPS 继续沿用原有只缩不放策略
  - PowerPoint 分支增加最小可读尺寸兜底
