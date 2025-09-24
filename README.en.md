<div align="center">

English | [简体中文](README.md)

</div>

[![Star History Chart](https://api.star-history.com/svg?repos=Achuan-2/SlideSCI&type=Date)](https://www.star-history.com/#Achuan-2/SlideSCI&Date)

Preview of Plugin Features

<img alt="PixPin_2025-08-29_15-10-51" src="https://s2.loli.net/2025/08/29/lRWKUwJCTjrk9ec.png" />

<img alt="PixPin_2025-08-29_15-11-12" src="https://s2.loli.net/2025/08/29/3dsS9UFtWL1niZx.png" />


## 📝 Development Background

Does anyone else share my long-standing grievances with PowerPoint?  😡:

💔 **No Image Titles**: Unlike Word, you can't directly add titles to images. You have to manually insert text boxes and struggle with alignment!

💔 **No Copy-Paste Element Positioning**: To keep similar elements in consistent positions across slides, you have to copy-paste and modify each time - no way to just copy-paste positions.

💔 **No Auto-Align for Images**: Insert multiple images and want them neatly arranged? Either drag each one manually for eternity or align them column by column.

💔 **No Code Block Insertion**: Have to copy-paste from external editors or screenshot code blocks - so tedious!

💔 **No LaTeX Formula Support**: Now I mainly rely on AI to recognize/generate math formulas in LaTeX format, which can't be directly pasted to PPT...

...

Most PPT plugins are packed with flashy but impractical features. As a graduate student doing weekly research progress reports, I need to quickly insert content and make clear presentations - aesthetics are secondary.

With AI's help, I developed solutions to these pain points quickly! (Over 99% of this plugin's code was AI-generated. Thank you AI teacher!)

In the spirit of open source, this plugin is publicly available on GitHub. Stars are appreciated!  🌟

GitHub: [https://github.com/Achuan-2/SlideSCI](https://github.com/Achuan-2/SlideSCI)

##  ✨ Key Features

- **Batch Add Image Titles:** Supports batch selection of images to add centered captions below them. Allows configuring auto-grouping of images and captions.
  <img alt="" src="https://s2.loli.net/2025/08/29/OoXlgpGdrtx2bEP.png" />

- **Batch Add Image Labels:** For scientific figures, supports label templates (`A`, `a`, `A)`, `a)`, `1`, `1)`). Default label font is `Arial`.

- **Auto-arrange Images:** Automatically aligns multiple images with configurable:
  - Sorting: By position or selection order
  - Layout: 
    - Column-max-width (for tabular layouts in academic figures)
    - Uniform height (uses first image's height by default)
    - Uniform width (waterfall flow, uses first image's width by default)
    - Custom spacing between columns/rows
  <img alt="" src="https://s2.loli.net/2025/08/29/RmxjZpTzGDL8evP.png" />

- **Copy-Paste Formatting:**
  - Style copying for shapes/text
  - Multi-element position copying (great for aligning elements across slides)
  - Bulk dimension pasting for uniform image sizes
  <img alt="" src="https://s2.loli.net/2025/08/29/q5vblI3nrDhewJ6.gif" />

- **Insert Syntax-Highlighted Code Blocks:**
  <img alt="" src="https://s2.loli.net/2025/08/29/jbSgDfnP69eZopV.png" />
  - Supported languages: MATLAB, Python, R, JavaScript, HTML, CSS, C#
  - Toggle between black/white background (default is black)

- **Insert LaTeX Math Formulas:**
  <img alt="" src="https://s2.loli.net/2025/08/29/qz9LMCuRB7AotDv.png" />

- **Insert Markdown Text:**
  <img alt="" src="https://s2.loli.net/2025/08/29/MPKOgWonijCsl4D.png" />
  - Preserves all formatting when pasting complete markdown documents
  - Inline formats: Bold, underline, superscript, subscript, italic, links, inline code/math
  - Block formats: 
    - Headings, lists (preserves hanging indents), code blocks (editable text boxes with syntax highlighting)
    - Tables (limited to 500px width with 1pt black borders by default)
    - Math formulas (editable text boxes)
    - Blockquotes (text boxes with black borders)
    - Task lists (converts to  ☑/☐ indicators)

* **Batch Add Image Labels**: For scientific figures, choose label templates (`A`, `a`, `A)`, `a)`). Default font is `Arial`.
    ![](https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/PixPin_2025-01-23_12-14-27-2025-01-23.png)

## 🪟 Supported Environments

- Developed on Windows 11 using [Visual Studio Tools for Office](https://www.visualstudio.com/vs/office-tools/) with C#
- Designed for Microsoft PowerPoint
- Compatible with WPS Office (Note: WPS version doesn't support LaTex formulas or Markdown insertion - may cause crashes)
- Windows only (Mac unsupported due to different plugin architectures)

## 🖥️ Installation

1. Download the plugin's `.exe` installer from GitHub [Releases](https://github.com/Achuan-2/my_ppt_plugin/releases)
2. Double-click to install
   
Important:
- Close PowerPoint before installation, otherwise the plugin won't load immediately

Required Dependencies (usually prompted automatically during installation):
1. [Microsoft .NET Framework 4.0+](https://www.microsoft.com/download/details.aspx?id=17718)
2. [Microsoft Visual Studio 2010 Tools for Office Runtime](https://www.microsoft.com/download/details.aspx?id=105522)

Troubleshooting:
- If the plugin doesn't appear in PowerPoint or shows "Runtime error loading COM add-in", install the dependencies above

## ❓ FAQs

* **How to add plugin features to the Quick Access Toolbar?**  
  Right-click a button and select "Add to Quick Access Toolbar."  
  ![](https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/PixPin_2025-01-16_16-56-07-2025-01-16.png)  
  Move the Quick Access Toolbar below the ribbon for easier access.

* **LaTeX formulas display incorrectly?**  
  Best for single-line formulas. For complex multi-line formulas, use [IguanaTex](https://github.com/Jonathan-LeRoux/IguanaTex).  
  See examples of PPT-specific LaTeX syntax [here](https://github.com/Achuan-2/my_ppt_plugin/issues/7).

## ❤️ Support My Work

If you like my plugin, please consider giving a star to the GitHub repository and making a donation. This will encourage me to continue improving this plugin.

![](https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/20241118182532-2024-11-18.png)

See the list of donors here: https://www.yuque.com/achuan-2


## 👨‍💻 Feedback

If you encounter any problems during use, you can provide feedback through the following ways:

1. Submit an [Issue](https://github.com/Achuan-2/my_ppt_plugin/issues) on GitHub
2. Send an email to: achuan-2@outlook.com


## 🔍 References & Acknowledgements

* [jph00/latex-ppt](https://github.com/jph00/latex-ppt): LaTeX in PowerPoint support
* [Markdig](https://github.com/xoofx/markdig): Markdown parsing support
* Thanks to Visual Studio Tools For Office for providing development support
* Thanks to the donors
* Thanks to all users who provided suggestions and feedback