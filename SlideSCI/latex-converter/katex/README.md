# Markdown 公式转换资源

Markdown 公式由 Jint 在插件进程内运行 KaTeX，输出 MathML 后交给 PowerPoint 导入为可编辑公式。无需 Node.js、pnpm、浏览器或 KaTeX 字体/CSS。独立的 LaTeX SVG 功能使用同样内置的 MathJax/Jint 流程。

- KaTeX 0.19.1：`katex.min.js` 来自[官方发行包](https://github.com/KaTeX/KaTeX/releases/tag/v0.19.1)，作为资源嵌入插件 DLL，许可见 `KaTeX-LICENSE.txt`。
- Jint 4.16.0、Acornima 1.7.0：通过 NuGet 引用 net462 程序集，随插件部署；许可见本目录的 LICENSE 和 NOTICE 文件。

后台预热会初始化共享引擎；转换串行执行，宏定义只在同一次转换请求内共享。KaTeX 支持范围与 MathJax 不完全相同，不支持的语法由现有 Markdown 流程保留源码并提示。

验证脚本：`tests/EmbeddedMathMl.Tests.ps1` 检查 .NET Framework 4.7.2、内嵌资源和清空 PATH 后的转换；`tests/MarkdownMathMl.Tests.ps1` 检查 Markdown 流程；`tests/MarkdownMathMl.PowerPoint.Tests.ps1` 检查 PowerPoint 导入后的原生公式结构。
