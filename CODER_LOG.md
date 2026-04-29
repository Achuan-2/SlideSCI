# Coder Log

用于记录当前协作过程中的代码与构建改动，不替代产品发布用的 `CHANGELOG.md`。

## 2026-04-29

### 1.4.6.38

- 升级版本号到 `1.4.6.38`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 按反馈继续放大「插入LaTeX SVG」三层取反线垂直间距
  - 将连续 overline 组阶梯步距从 `Math.max(120, thickness * 3.2)` 翻倍为 `Math.max(240, thickness * 6.4)`
  - 保留 `1.4.6.37` 的最终横线坐标扫描与 `viewBox` 自适应扩展，避免最高线被 SVG 裁掉
- 转换验证
  - `\overline{A}\overline{B}\overline{A}B` 输出 3 条横线，y 坐标约为 `-1413/-1173/-933`
  - 根 SVG `viewBox` 为 `0 -1495.64 3018 1578.28`，最高横线仍在可见范围内
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.38.exe`
- 轻量验证
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.38`，日志显示 `Installation process succeeded`
  - 安装后 `{app}\latex-converter\latex-to-svg.js` 包含 `Math.max(240, thickness * 6.4)` 与 `rectBounds`
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.38`

### 1.4.6.37

- 升级版本号到 `1.4.6.37`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复「插入LaTeX SVG」三层取反线最高线被裁剪的问题
  - 根因：`1.4.6.36` 把第一条 overline 上移后，根 SVG `viewBox` 仍按原始 MathJax 高度扩展，PowerPoint/WPS 导入时会裁掉超出 viewBox 的最高线
  - `addRootViewBoxPadding` 改为扫描最终 SVG 中的矩形横线坐标，按最终 `rectBounds` 重新扩展 viewBox
  - 继续保留三条高度不一的取反线：第一条覆盖第 1-3 个变量，第二条覆盖第 2-3 个变量，第三条只覆盖第 3 个变量
- 转换验证
  - `\overline{A}\overline{B}\overline{A}B` 输出 3 条横线，y 坐标约为 `-1173/-1053/-933`
  - 根 SVG `viewBox` 从 `0 -1115.64 3018 1198.28` 扩为 `0 -1255.64 3018 1338.28`，最高横线已在可见范围内
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.37.exe`
- 轻量验证
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.37`，日志显示 `Installation process succeeded`
  - 安装后 `{app}\latex-converter\latex-to-svg.js` 长度为 `25809`，包含 `targetEndIndex = index === 0` 与 `rectBounds`
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.37`

### 1.4.6.36

- 升级版本号到 `1.4.6.36`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 继续优化「插入LaTeX SVG」连续取反横线覆盖范围
  - 针对 `\overline{A}\overline{B}\overline{A}B` 这类连续单变量 overline 组，改为按用户确认的 LaTeX 风格覆盖：
    - 第一条横线覆盖第 1-3 个变量
    - 第二条横线覆盖第 2-3 个变量
    - 第三条横线只覆盖第 3 个变量
  - 将阶梯垂直间距从 `Math.max(70, thickness * 1.8)` 增加到 `Math.max(120, thickness * 3.2)`，避免三层高度差过小
- 转换验证
  - `\overline{A}\overline{B}\overline{A}B` 输出单层 SVG，无嵌套 SVG
  - 3 条横线坐标约为：
    - `x=170-2089, y=-1173`
    - `x=920-2089, y=-1053`
    - `x=1679-2089, y=-933`
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.36.exe`
- 轻量验证
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.36`，日志显示 `Installation process succeeded`
  - 安装后 `{app}\latex-converter\latex-to-svg.js` 长度为 `24469`，包含 `targetEndIndex = index === 0` 与 `Math.max(120, thickness * 3.2)`
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.36`

### 1.4.6.35

- 升级版本号到 `1.4.6.35`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 继续优化「插入LaTeX SVG」连续取反横线视觉
  - `staggerConsecutiveOverlineBars` 不再只做 y 方向错开
  - 对连续 3 条及以上 overline 横线组，前面的横线会扩展为覆盖相邻变量窗口
  - `\overline{A}\overline{B}\overline{A}B` 转换后：
    - 第一条横线覆盖第 1-2 个变量
    - 第二条横线覆盖第 2-3 个变量
    - 第三条横线覆盖第 3 个变量
    - 同时保持阶梯 y 位置
- 转换验证
  - `\overline{A}\overline{B}\overline{A}B` 输出单层 SVG，无嵌套 SVG
  - 3 条横线宽度分别约为 `1169/1169/410`，位置为阶梯分布
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.35.exe`
- 轻量验证
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.35`，日志显示 `Installation process succeeded`
  - 安装后 `{app}\latex-converter\latex-to-svg.js` 长度为 `24365`，包含当前版本的连续取反覆盖窗口逻辑
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.35`

### 1.4.6.34

- 升级版本号到 `1.4.6.34`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复「插入LaTeX SVG」实际加载旧版 `latex-to-svg.js` 的问题
  - 根因：安装器只同步 DLL/config/manifest 到 `{app}` 根目录，没有同步当前版本的 `latex-converter`
  - `LatexToSvgConverter.cs` 调整脚本搜索顺序，优先使用 `AppDomain.CurrentDomain.BaseDirectory` 和程序集目录下的 `latex-converter`，再兼容固定安装目录和 AppData
  - `installer/SlideSCI.iss` 新增 `CopyLatexConverterFilesToAppRoot`，安装后把当前版本 `Application Files\SlideSCICompat_x_y_z\latex-converter` 下的 `latex-to-svg.js`、`package.json`、`package-lock.json` 同步到 `{app}\latex-converter`
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.34.exe`
- 轻量验证
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.34`，日志显示 `Installation process succeeded`
  - 安装后 `D:\SlideSCI_WPS_PowerPoint_Compat\latex-converter\latex-to-svg.js` 长度为 `24037`，包含 `staggerConsecutiveOverlineBars`
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.34`

### 1.4.6.33

- 升级版本号到 `1.4.6.33`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 优化「插入LaTeX SVG」连续单变量取反线的垂直层次
  - 新增 `staggerConsecutiveOverlineBars`，在 SVG 后处理中识别相邻很近的 overline 横线组
  - 对 3 条及以上的连续横线组做阶梯式 y 偏移，保留“第一条最高、第二条中间、第三条最低”的 LaTeX 视觉
  - 间隔较远的独立取反项不参与分组，避免影响 `\overline{A}C\overline{A}B` 这类普通表达式
- 转换验证
  - `\overline{A}\overline{B}\overline{A}B` 输出单层 SVG，无嵌套 SVG
  - 该样例 3 条横线 y 位置为阶梯分布
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.33.exe`
- 轻量验证
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.33`，日志显示 `Installation process succeeded`
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.33`

### 1.4.6.32

- 升级版本号到 `1.4.6.32`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 调整「插入LaTeX SVG」单变量取反线策略
  - 撤销 `1.4.6.31` 中将 `\overline{A}` 归一化为 `\bar{A}` 的处理，避免改变 MathJax/LaTeX 的原始排版模型
  - 保留 `\overline{A}` 原始语义，让 A/B 的取反线保留自然高度差
  - 将 overline 横线内缩从 `18%/120` 调整为 `24%/170`，使横线更接近字母宽度并减少相邻重叠
- 转换验证
  - `=\overline{A}\overline{B}\overline{A}B+\overline{A}C\overline{A}B` 输出单层 SVG，无嵌套 SVG
  - 该样例中取反线宽度约从 510 缩至 410，A/B 横线保留不同 y 位置
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.32.exe`
- 轻量验证
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.32`，日志显示 `Installation process succeeded`
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.32`

### 1.4.6.31

- 升级版本号到 `1.4.6.31`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 优化「插入LaTeX SVG」中单字母取反横线宽度/对齐问题
  - 在 `latex-converter/latex-to-svg.js` 中新增 `normalizeSingleTokenOverline`
  - 将单字符 `\overline{A}`、`\overline{B}` 这类布尔变量取反归一化为 `\bar{A}`、`\bar{B}`
  - 多字符和嵌套表达式仍保留 `\overline{...}`，例如 `\overline{AB}`、`\overline{\bar{A}B}`
  - 目标是让单变量取反线跟字母宽度更接近，避免 MathJax 的 token 盒子宽度导致横线过宽或相邻重合
- 转换验证
  - `=\overline{A}\overline{B}\overline{A}B+\overline{A}C\overline{A}B` 输出单层 SVG，无嵌套 SVG
  - `\overline{\overline{A}B}+...` 输出单层 SVG，无嵌套 SVG
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.31.exe`
- 轻量验证
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.31`，日志显示 `Installation process succeeded`
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.31`

### 1.4.6.30

- 升级版本号到 `1.4.6.30`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复 Markdown 插入时复杂文档只剩表格、普通段落/列表/加粗丢失的问题
  - 收紧 `SplitMarkdownIntoSegments` 的表格正则，避免表格分隔行中的 `\s` 跨行吞掉后续 Markdown
  - 修复引用块正则允许空匹配的问题，避免分段器产生无意义匹配
  - 用用户提供的 Markdown 结构验证分段结果为 `普通文本 -> 表格 -> 普通文本`
- 继续优化「插入LaTeX SVG」连续取反横线
  - 将 `\overline` 横线内缩从 `8%/28` 调整为 `18%/120`
  - `=\overline{A}\overline{B}\overline{A}B` 转换后 3 条横线间距明显拉开，仍保持单层 SVG
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.30.exe`
- 轻量验证
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.30`，日志显示 `Installation process succeeded`
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.30`

### 1.4.6.29

- 升级版本号到 `1.4.6.29`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复 Markdown 插入 `$$...$$` 公式块时偶发 `Shape (unknown member): Object does not exist` 的问题
  - `InsertMathBlock` 原来返回的是用于触发公式插入流程的旧 `equationShape`
  - 改为返回真正写入公式内容并执行 `EquationProfessional` 后的 `equationShape2`
  - Markdown 排版后续读取 `shape.Height` 时不再访问失效 COM 对象
  - Markdown 公式块也复用 `ConvertLatexToPowerPointEquationText`，保持 `\xrightarrow{}` 等转换一致
- 继续优化「插入LaTeX SVG」的取反横线
  - 修正嵌套 SVG 扁平化时对内部 viewBox 偏移的双重平移，避免横线位置整体偏移
  - 对 MathJax overline 使用的 `data-c="2013"` 横线做轻微左右内缩，避免相邻 `\overline{A}\overline{B}\overline{A}` 被 Office/WPS 渲染成一条连续长线
- 转换验证
  - `=\overline{A}\overline{B}\overline{A}B` 输出 3 条独立取反横线，且只有 1 个 `<svg>`
  - 用户提供的完整公式样例输出只有 1 个 `<svg>`，无嵌套 SVG
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.29.exe`
- 轻量验证
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.29`，日志显示 `Installation process succeeded`
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.29`

### 1.4.6.28

- 升级版本号到 `1.4.6.28`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 优化「插入LaTeX SVG」中上划线/取反嵌套和高度裁剪问题
  - `latex-converter/latex-to-svg.js` 改为输出 Office/WPS 更稳定的单层 SVG
  - 将 MathJax 用于 `\overline`、`\overbrace`、扩展箭头等结构的内部 `<svg>` 扁平化到同一个坐标系
  - 根 SVG 增加 `overflow="visible"`，并给 `viewBox` 上下增加安全边距，降低上划线/重音被裁剪的概率
- 转换验证
  - `=\overline{A}\overline{B}\overline{A}B` 输出只有 1 个 `<svg>`，无嵌套 SVG
  - 用户提供的完整公式样例输出只有 1 个 `<svg>`，无嵌套 SVG
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.28.exe`
- 轻量验证
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.28`，日志显示 `Installation process succeeded`
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.28`

### 1.4.6.27

- 升级版本号到 `1.4.6.27`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复 PowerPoint「插入LaTeX文字」仍会因为 `\xrightarrow{}` / `\xleftarrow{}` 回退 SVG 的问题
  - PowerPoint 下该按钮现在只在 WPS 宿主时走 SVG；PowerPoint 宿主始终插入可编辑公式
  - 新增 `ConvertLatexToPowerPointEquationText`，在写入 PowerPoint 原生公式前将 `\xrightarrow{abc}` 转为 Office 公式编辑器可解析的 `->\above(abc)`
  - 同理将 `\xleftarrow{abc}` 转为 `<-\above(abc)`
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.27.exe`
- 轻量验证
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.27`，日志显示 `Installation process succeeded`
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.27`

### 1.4.6.26

- 升级版本号到 `1.4.6.26`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复 PowerPoint「插入LaTeX文字」中 `\xrightarrow{}` / `\xleftarrow{}` 原样显示的问题
  - PowerPoint 原生公式不支持这类 amsmath 扩展箭头命令，保留原生路径会显示为 `\xrightarrow` 文本
  - 新增 `RequiresSvgForPowerPointLatex`，仅对 `\xrightarrow` / `\xleftarrow` 自动回退 SVG
  - 普通 LaTeX 仍走 PowerPoint 原生公式文本路径，避免再次把「插入LaTeX文字」整体变成 SVG
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.26.exe`
- 轻量验证
  - MathJax 转换 `\xrightarrow{abc}` 输出 SVG 成功
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.26`，日志显示 `Installation process succeeded`
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.26`

### 1.4.6.25

- 升级版本号到 `1.4.6.25`
  - `SlideSCI/SlideSCI.csproj`
  - `SlideSCI/Properties/AssemblyInfo.cs`
- 修复 PowerPoint 插入 Markdown 报 `System.Memory, Version=4.0.1.2` 找不到的问题
  - `packages.config` 已声明 `System.Memory`、`System.Buffers`、`System.Numerics.Vectors`，但 `SlideSCI.csproj` 未引用这些运行时 DLL，VSTO 发布目录不会带出依赖
  - 在项目文件中补齐 `System.Memory.dll`、`System.Buffers.dll`、`System.Numerics.Vectors.dll` 引用，并显式 `Private=True`
  - 同步将 `System.Runtime.CompilerServices.Unsafe.dll` 显式设为 `Private=True`
- 修复 PowerPoint「插入LaTeX文字」被复杂公式判断改走 SVG 的问题
  - PowerPoint 下始终走原生公式文本插入路径
  - WPS 下仍保留 SVG 兼容路径，避免调用 PowerPoint 专用公式命令
- 重新编译并生成安装包
  - `artifacts/dist/SlideSCI_WPS_PowerPoint_Compat_v1.4.6.25.exe`
- 安装验证
  - 发布目录包含 `System.Memory.dll.deploy`、`System.Buffers.dll.deploy`、`System.Numerics.Vectors.dll.deploy`
  - 使用 `/VERYSILENT /LOG` 执行 `1.4.6.25`，日志显示 `Installation process succeeded`
  - 等待安装器进程退出后重复静默安装 2 次，均 `ExitCode=0`，日志无 fatal/uninstall rollback
  - 注册表显示 PowerPoint `LoadBehavior=3`
  - 注册表显示当前安装版本为 `SlideSCI WPS PowerPoint version 1.4.6.25`
  - 安装根目录存在 `System.Memory.dll`、`System.Buffers.dll`、`System.Numerics.Vectors.dll`、`System.Runtime.CompilerServices.Unsafe.dll`
  - 安装后无残留 `SlideSCI_WPS_PowerPoint_Compat_v*` 安装器进程

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
