# SlideBridge Office

English | [简体中文](README.md)

SlideBridge Office is a Windows VSTO add-in for Microsoft PowerPoint and WPS Presentation. It focuses on common slide-editing workflows for research reports and technical presentations: figure layout, format reuse, code blocks, Markdown insertion, and LaTeX formulas.

This project is a recreation of the original author's project and is maintained here with additional WPS and PowerPoint compatibility work.

## Features

- Batch image captions: add centered captions to selected images, with optional grouping.
- Batch figure labels: add labels such as `A`, `a`, `A)`, `a)`, `1`, and `1)`.
- Image auto layout: arrange images by position or selection order, with column-width, uniform-height, and waterfall-style layout modes.
- Format reuse: copy and paste shape/text formatting, positions, dimensions, and crop settings.
- Code blocks: insert syntax-highlighted code blocks with black or white background.
- LaTeX formulas: insert formulas through native PowerPoint equations or SVG rendering.
- Markdown insertion: insert text, headings, lists, tables, code blocks, blockquotes, and inline math.

## Supported Environments

- Microsoft PowerPoint for Windows.
- WPS Presentation for Windows.
- Runtime dependencies:
  - Microsoft .NET Framework 4.7.2 or later.
  - Microsoft Visual Studio 2010 Tools for Office Runtime.
  - Node.js for LaTeX SVG rendering. Run `npm install` inside the installed `latex-converter` directory before using that feature.

## Installation

Download the installer from this project's Release page, close PowerPoint/WPS, then run the installer. After installation, reopen PowerPoint or WPS and use the `SlideBridge` ribbon tab.

If the add-in does not appear, check:

- PowerPoint or WPS was closed during installation.
- VSTO Runtime is installed.
- The add-in is not disabled in PowerPoint COM Add-ins.
- WPS add-in registry entries were not blocked or overwritten by stale state.

## LaTeX SVG Setup

The `Insert LaTeX SVG` feature uses local Node.js and MathJax:

1. Install Node.js.
2. Open the `latex-converter` directory in the installed plugin folder.
3. Run `npm install`.
4. Restart PowerPoint or WPS.

## Feedback

- GitHub: [https://github.com/jacywallny/](https://github.com/jacywallny/)
- Email: [jacywalln@gmail.com](mailto:jacywalln@gmail.com)

## References

- [jph00/latex-ppt](https://github.com/jph00/latex-ppt): LaTeX in PowerPoint support.
- [Markdig](https://github.com/xoofx/markdig): Markdown parsing.
- [MathJax](https://github.com/mathjax/MathJax): LaTeX to SVG rendering.
- Visual Studio Tools for Office: PowerPoint add-in development.

## Notice

The scripts and add-in code in this project are intended for learning, research, and personal productivity workflows. Please evaluate applicability and risks before use. This project is licensed under `AGPL-3.0`.
