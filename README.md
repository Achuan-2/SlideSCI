<div align="center">

[English](README.en.md) | 简体中文

</div>


[![Star History Chart](https://api.star-history.com/svg?repos=Achuan-2/SlideSCI&type=Date)](https://www.star-history.com/#Achuan-2/SlideSCI&Date)

插件功能预览

<img alt="PixPin_2025-08-29_15-10-51" src="https://s2.loli.net/2025/08/29/lRWKUwJCTjrk9ec.png" />

<img alt="PixPin_2025-08-29_15-11-12" src="https://s2.loli.net/2025/08/29/3dsS9UFtWL1niZx.png" />




## 1 ❤️ 友情推荐

我的Adobe illustrator插件：[https://github.com/Achuan-2/illustrator_sci_plugin](https://github.com/Achuan-2/illustrator_sci_plugin)，专为科研组图设计，支持复制粘贴相对位置、形状尺寸批量设置、图片一键自动排列，一键添加子图label

<img alt="image" src="https://s2.loli.net/2025/08/29/OxzyYpbiWrPKgfG.png" style="width: 249px;" />​

## 2 📝 开发背景

不知道有没有人和我一样，对PPT积怨已久😡：

💔​**不能添加图片标题**：图片没法像word一样直接添加图片标题，只能手动插入文本框，对齐半天还歪七扭八！

💔​**不能复制元素的位置粘贴给另一个元素**：不同页PPT的类似元素要保持同一个位置，只能复制粘贴再修改，无法直接复制粘贴位置

💔**不能图片自动排列整齐：**  插入多张图片后，想要多行多列整齐排列？要么一张一张手动拖动，对齐到天荒地老！要么先一列列水平对齐再垂直对齐

💔**不能插入代码块：**  只能从外部编辑器（如VSCode）或专门网站复制粘贴，或者截图、生成图片粘贴代码块，有点麻烦

💔**不能插入latex数学公式：**  现在我基本上靠ai来识别和生成数学公式，公式都是latex数学公式格式，不方便直接粘贴到PPT里

……

市面上的ppt插件花里胡哨的功能一大堆，没几个能用得上。对我而言，每周要做研究生科研进展工作汇报，要的就是快速插入内容、做出内容清晰的PPT，不追求太美观。

在AI的帮助下，很快就把这些痛点功能都开发出来了 ！真的成就感满满！（这个插件99%的代码都是ai生成的，感谢AI老师！）

本着开源的精神，这个插件也在Github上开源了，欢迎大家给我点小星星！

Github地址：[https://github.com/Achuan-2/SlideSCI](https://github.com/Achuan-2/SlideSCI)

## 3 ✨ 主要功能

- **批量添加图片标题**：支持批量选中图片后，批量在图片下方添加居中图题，支持设置图片和图题是否自动编组  
  ​<img alt="" src="https://s2.loli.net/2025/08/29/OoXlgpGdrtx2bEP.png" />
- **批量添加图片标签**：用于科研绘图，可选择添加的标签模板（`A`、`a`、`A)`、`a)`、`1`、`1)`），默认添加的标签字体为`Arial`<img alt="" src="assets/network-asset-68747470733a2f2f666173746c792e6a7364656c6976722e6e65742f67682f41636875616e2d322f5069634265642f6173736574732f50697850696e5f323032352d30312d32345f31382d33342d31312d323032352d30312d32342e706e6-20250417103743-vry4kvj.png" />
- **图片自动排列**：可以自动排列多张图片，支持设置排序方式、排列方式、每列多少张图片、列间距多少、行间距多少（默认为空，为列间距大小）、图片宽高

  - **排序方式**：

    - 根据位置排序：可以根据用户手动粗排的图片位置，来自动判断图片的排列顺序。
    - 根据多选顺序：根据用户的多选选择顺序来排列
  - **排列方式**：

    - **列最大宽度占位排列**：按每列的最大宽度来占位排列，以保持表格布局，适用于科研绘图的图片排版
    - **统一高度排列**：默认统一图片的高度进行排列，如果不设置图片高度，则使用第一张图片的高度来统一图片高度
    - **统一宽度瀑布流**：默认统一图片的宽度进行排列，如果不设置图片宽度，则使用第一张图片的宽度来统一图片宽度  
      ​<img alt="" src="https://s2.loli.net/2025/08/29/RmxjZpTzGDL8evP.png" />
- **复制粘贴格式功能**  
  ​<img alt="" src="https://s2.loli.net/2025/08/29/3fjcw5KWbsLJkAo.png" />

  - **复制粘贴格式**：可以复制形状、文字的格式，粘贴给其他元素。
  - **复制位置和粘贴元素位置**：可以复制多个元素的位置，粘贴给其他元素（可以多选复制和粘贴！），可以用来让不同页的PPT的多个元素位置一致，或者让同一页的不同元素都是一个位置（可以先排好一组元素，用这个功能让另一组元素自动排好，再调整位置）。  
    ​<img alt="" src="https://s2.loli.net/2025/08/29/q5vblI3nrDhewJ6.gif" />  
      
    ​<img alt="" src="https://s2.loli.net/2025/08/29/u19w8IMkOYjE3rz.gif" />
  - **复制和粘贴元素宽高**：支持多选同时粘贴宽高，快速统一图片宽高
- **支持插入代码块自动高亮**  
  ​<img alt="" src="https://s2.loli.net/2025/08/29/jbSgDfnP69eZopV.png" />

  - **支持代码语言高亮列表**：matlab、python、r、js、html、css、csharp、Fortran
  - **支持切换黑白背景色**：默认是黑色背景色，切换为背景色，只需要点击「代码黑色背景色」按钮取消激活状态即可
- **支持插入latex数学公式**  
  ​「插入Markdown」可以快速插入文字与数学混排
  <img width="1926" height="1106" alt="image" src="https://github.com/user-attachments/assets/8d345baf-41af-473d-a182-3c7d2864c155" />

  「插入LaTeX文字」单独插入可编辑的数学公式

  <img width="1889" height="1203" alt="image" src="https://github.com/user-attachments/assets/d33eba81-4241-4ec6-bc73-057087be8a5e" />

  如果数学公式比较复杂，PPT原生不支持，可以选择「插入LaTeX svg」，支持更多latex公式样式，只需要简单配置下node.js环境即可使用，而不需要像IguanaTeX等插件除了安装LaTeX环境还要安装很多插件才能实现latex转svg，本插件的安装简单很多
  <img width="1756" height="1202" alt="PixPin_2025-09-30_10-21-00" src="https://github.com/user-attachments/assets/4d7732b2-47da-41e1-bb22-ddfceab4604b" />
  插入的svg的图片提示文字会保留原始LaTeX代码，方便修改后再插入
  <img width="1768" height="979" alt="PixPin_2025-09-30_10-25-48" src="https://github.com/user-attachments/assets/b5bbeaf1-8b12-44b1-9664-4591741e83ea" />



  > 配置「插入LaTeX svg」功能的方法
  >
  > - 配置[Node.js](https://nodejs.org/zh-cn/)环境
  > - 进入插件安装文件夹的`latex-converter`文件夹，默认为`%APPDATA%\Achuan-2\SlideSCI\latex-converter`
  > - 打开命令行，运行`npm install`安装环境
  > - 配置完成，可以使用「插入LaTeX svg」功能


- **支持插入Markdown文本**：可以直接把markdown整篇笔记一口气全部粘贴到PPT里！并按原文顺序排列！  
  ​<img alt="" src="https://s2.loli.net/2025/08/29/MPKOgWonijCsl4D.png" />

  - 支持行内格式：加粗、下划线、上标、下标、斜体、链接、行级代码、行级数学公式
  - 支持块级格式：标题、列表、代码块、表格、数学公式、引述块

    - 列表特别处理：

      - 保留列表悬挂缩进：常规粘贴html中的列表到PPT里会丢失悬挂缩进，本插件粘贴列表可以保留悬挂缩进
      - 支持任务列表粘贴，列表项符号会转化为☑和☐，代表完成和未完成
    - 代码块特别处理：

      - 独立文本框，支持设置黑白背景高亮配色，支持PPT直接再编辑
    - 表格特别处理：

      - 默认限制500px宽度，添加1pt黑色边框
    - 数学公式特别处理：

      - 独立文本框，支持PPT直接再编辑
    - 引述块特别处理：

      - 独立文本框，添加黑色边框

## 4 🪟 支持环境

插件在Windows11 使用[Visual Studio Tools For Office](https://www.visualstudio.com/de/vs/office-tools/) 和C#语言开发，专为Microsoft Powerpoint设计，兼容安装到WPS（注：WPS版本不支持插入latex数学公式、插入markdown笔记，强行使用会造成卡死）

> 插件只支持Windows端，不支持Mac端（因为Windows端的PPT插件和Mac端插件开发架构不一样，市面上很多的PPT插件，Windows端和Mac端都是独立开发的界面）

## 5 🖥️ 安装方法

下载本插件Github页面[Release](https://github.com/Achuan-2/my_ppt_plugin/releases)中的exe文件，双击安装即可

注意：安装时需要先退出PPT，否则PPT不会即时加载该插件

需要的安装环境（一般插件安装时会自动提示安装）

- 运行环境1：[Microsoft .Net Framework 4.0或更高](https://www.microsoft.com/zh-cn/download/details.aspx?id=17718)
- 运行环境2：[Microsoft Visual Studio 2010 Tools for Office Runtime](https://www.microsoft.com/zh-cn/download/details.aspx?id=105522)

如果插件安装后无法正常运行、没有在PPT中显示，或者在开发工具→COM加载项里勾选插件，提示“未加载。加载COM加载项时出现运行错误”，请尝试安装上面的环境依赖。

## 6 📝 使用介绍视频

- [让AI帮我我开发了一个PPT插件！支持一键添加图片标题，复制粘贴位置、一键图片对齐、一键插入代码块、一键插入latex公式！_哔哩哔哩_bilibili](https://www.bilibili.com/video/BV15YcmeSEbq/?spm_id_from=0.0.homepage.video_card.click&vd_source=b4a1fcb6dce305e26d8d16d9cbb71304)
- [可能是全网第一个支持插入Markdown到PPT的PPT插件_哔哩哔哩_bilibili](https://www.bilibili.com/video/BV1VXcZe2EyK/?spm_id_from=0.0.homepage.video_card.click)

## 7 🤝参与开发

我是如何开发PPT插件的分享：[如何开发一个PPT插件：使用VSTO开发](https://mp.weixin.qq.com/s/LpOw5tdbHTbC1FsG4seEOw)

## 8 ❓ 常见问题

- **如何把插件的功能添加到PPT的快捷工具栏？**   
  按钮右击，添加到「快速访问工具栏」即可  
  ​<img alt="" src="https://s2.loli.net/2025/08/29/RdxoUpzIcZMTv6u.png" />  
  可以把「快速访问工具栏」放在下方，更方便使用

- **插入latex公式，显示不正常？**

  - 插件比较适合插入单行数学公式，对于复杂的多行语法，推荐使用IguanaTex插件
  - PPT特殊latex语法举例见：[#7](https://github.com/Achuan-2/SlideSCI/issues/7)

## 9 ❤️ 用爱发电

开发和维护不易，如果喜欢我的插件，欢迎给GitHub仓库点star和捐赠，这会激励我继续完善此插件。

<img alt="" src="https://s2.loli.net/2025/08/29/tI4HCGzql17nr2D.png" />

## 10 👨💻 问题反馈

如果使用过程中遇到问题，可以通过以下方式反馈:

1. 在 GitHub 上提 [Issue](https://github.com/Achuan-2/my_ppt_plugin/issues)
2. 发送邮件到: [achuan-2@outlook.com](mailto:achuan-2@outlook.com)

## 11 🔍 参考项目与致谢

- [jph00/latex-ppt](https://github.com/jph00/latex-ppt): LaTeX in PowerPoint 支持
- [Markdig](https://github.com/xoofx/markdig): Markdown 解析支持
- [MathJax](https://github.com/mathjax/MathJax): 数学公式转SVG支持
- 感谢 Visual Studio Tools For Office 提供的开发支持
- 感谢所有提供建议反馈以及捐赠赞赏的用户

## 12 📄 特别说明

- 本仓库发布的`SlideSCI`项目中涉及的任何脚本，仅用于测试和学习研究，禁止用于商业用途，不能保证其合法性，准确性，完整性和有效性，请根据情况自行判断。
- ​`本人` 对任何脚本问题概不负责，包括但不限于由任何脚本错误导致的任何损失或损害.
- 未经授权，请勿将`SlideSCI`项目的任何内容用于商业或非法目的，否则后果自负。
- 以任何方式查看此项目的人或直接或间接使用`SlideSCI`项目的任何脚本的使用者都应仔细阅读此声明。`本人` 保留随时更改或补充此免责声明的权利。一旦使用并复制了任何相关脚本或`SlideSCI`项目，则视为您已接受此免责声明。
- 本项目遵循`AGPL-3.0 License`协议，如果本特别声明与`AGPL-3.0 License`协议有冲突之处，以本特别声明为准。



> 您使用或者复制了本仓库且本人制作的任何代码或项目，则视为`已接受`此声明，请仔细阅读  
> 您在本声明未发出之时点使用或者复制了本仓库且本人制作的任何代码或项目且此时还在使用，则视为`已接受`此声明，请仔细阅读
