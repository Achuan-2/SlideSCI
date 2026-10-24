# 内置 SVG 转换

`mathjax-svg.min.js` 由 MathJax 3.2.2 的 TeX、SVG 和 liteAdaptor 模块打包，包含 SVG 字形路径，在 Jint 中运行。插件用户无需 Node.js、pnpm、浏览器、网络或 LaTeX 环境。引擎在首次使用时初始化，后续共享引擎；每次转换新建解析器，避免宏与公式编号串入下一次调用。

开发者更新资源时，在上级目录运行 `pnpm install` 和 `pnpm run build:svg`。开发工具使用 Node.js，生成资源已提交并嵌入插件 DLL，普通插件构建无需重新打包 JavaScript。

许可证见 `MathJax-LICENSE.txt`。生成脚本尾部保留了 MathJax、mhchem 等模块的版权声明。更新资源后运行 `tests/EmbeddedMathMl.Tests.ps1`，再用 Windows PowerShell -STA 运行 `tests/EmbeddedSvg.PowerPoint.Tests.ps1`。

上级目录的 `latex-to-svg.js` 仅保留为开发时的 Node.js 对照工具，不再部署或供插件调用。
