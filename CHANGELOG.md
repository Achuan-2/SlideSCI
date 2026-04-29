# Changelog

## v1.4.6.52 / 2026-04-29

- Rename the public-facing project to `SlideBridge Office`.
- Update public contact information to:
  - GitHub: `https://github.com/jacywallny/`
  - Email: `jacywalln@gmail.com`
- Remove public original-author links, video links, and funding material.
- Keep internal VSTO compatibility identifiers unchanged for PowerPoint/WPS upgrade stability.

## v1.4.6.51 / 2026-04-29

- Fix PowerPoint insertion of long Chinese text mixed with inline LaTeX formulas.
- Mixed text with `$...$` is inserted as a wrapping text box instead of a single long SVG.
- Pure formula input continues to use the SVG path.

## v1.4.6.50 / 2026-04-29

- Improve PowerPoint LaTeX SVG sizing for long formulas.
- Keep WPS sizing behavior unchanged.

## v1.4

- Add WPS and PowerPoint compatible installer flow.
- Support image captions, image labels, image auto layout, format copy/paste, Markdown insertion, code blocks, and LaTeX SVG formulas.
- Improve repeated install behavior and add-in load reliability.
