# SlideBridge Office

[English](README.en.md) | 简体中文

SlideBridge Office 是一个面向 Windows 端 Microsoft PowerPoint 与 WPS 演示的 VSTO 插件，聚焦科研汇报和日常幻灯片编辑中的高频操作：图片排版、格式复用、代码块插入、Markdown 粘贴和 LaTeX 公式插入。

本项目复刻于原作者项目，并在此基础上做 WPS 与 PowerPoint 兼容维护。

## 主要功能

- 批量添加图片标题：支持给多张图片批量添加居中图题，并可选择是否与图片自动编组。
- 批量添加图片标签：适合科研组图，可使用 `A`、`a`、`A)`、`a)`、`1`、`1)` 等标签模板。
- 图片自动排列：支持按位置或选择顺序排序，支持列最大宽度占位、统一高度、统一宽度瀑布流等排版方式。
- 复制粘贴格式：支持复制形状和文字格式、复制粘贴元素位置、批量粘贴宽高和裁剪参数。
- 插入代码块：支持语法高亮，并可切换黑色或白色背景。
- 插入 LaTeX 公式：支持 PowerPoint 原生公式和 SVG 公式两种路径。
- 插入 Markdown：支持普通文本、标题、列表、表格、代码块、引用块和行内数学公式。

## 支持环境

- Windows 端 Microsoft PowerPoint。
- Windows 端 WPS 演示。
- 运行环境依赖：
  - Microsoft .NET Framework 4.7.2 或更高版本。
  - Microsoft Visual Studio 2010 Tools for Office Runtime。
  - 使用 LaTeX SVG 功能时需要 Node.js，并在安装目录的 `latex-converter` 文件夹中执行 `npm install`。

## 安装与使用

从本项目 Release 页面下载安装包，关闭 PowerPoint/WPS 后运行安装程序。安装完成后重新打开 PowerPoint 或 WPS，功能会出现在 `SlideBridge` 选项卡中。

如果插件没有显示，请检查：

- PowerPoint 或 WPS 是否在安装时仍处于打开状态。
- VSTO Runtime 是否已安装。
- PowerPoint COM 加载项中插件是否被禁用。
- WPS 的加载项注册项是否被安全软件或旧版本残留覆盖。

## LaTeX SVG 配置

`插入 LaTeX SVG` 使用本地 Node.js + MathJax 转换公式：

1. 安装 Node.js。
2. 进入插件安装目录下的 `latex-converter` 文件夹。
3. 执行 `npm install`。
4. 重新打开 PowerPoint 或 WPS 后使用该功能。

## 反馈

- GitHub: [https://github.com/jacywallny/](https://github.com/jacywallny/)
- Email: [jacywalln@gmail.com](mailto:jacywalln@gmail.com)

## 参考与致谢

- [jph00/latex-ppt](https://github.com/jph00/latex-ppt): LaTeX in PowerPoint 支持。
- [Markdig](https://github.com/xoofx/markdig): Markdown 解析。
- [MathJax](https://github.com/mathjax/MathJax): LaTeX 到 SVG 转换。
- Visual Studio Tools for Office: PowerPoint 加载项开发基础。

## 说明

本项目中的脚本和插件代码仅用于学习、研究和个人效率改进，请自行判断适用场景与风险。本项目遵循 `AGPL-3.0 License` 协议。
