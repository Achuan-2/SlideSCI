// Development entry point. The bundled result runs in Jint without Node or a DOM.
const { mathjax } = require('mathjax-full/js/mathjax.js');
const { TeX } = require('mathjax-full/js/input/tex.js');
const { AllPackages } = require('mathjax-full/js/input/tex/AllPackages.js');
const { SVG } = require('mathjax-full/js/output/svg.js');
const { liteAdaptor } = require('mathjax-full/js/adaptors/liteAdaptor.js');
const { RegisterHTMLHandler } = require('mathjax-full/js/handlers/html.js');

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);

globalThis.slideSciToSvg = function (source, display) {
    // A new parser keeps definitions and equation numbers local to this request.
    const document = mathjax.document('', {
        InputJax: new TeX({ packages: AllPackages.filter(name => name !== 'noerrors' && name !== 'noundefined') }),
        OutputJax: new SVG({ fontCache: 'none' })
    });
    const node = document.convert(source, { display, em: 20, ex: 10, containerWidth: 1600 });
    // Standalone SVGs cannot rely on MathJax's page-level CSS. Inline the SVG
    // presentation rules, including table/proof-tree lines that otherwise vanish.
    for (const element of adaptor.tags(node, 'path')) {
        if (adaptor.hasAttribute(element, 'data-c')) adaptor.setAttribute(element, 'stroke-width', '3');
    }
    for (const tag of ['line', 'rect']) {
        for (const element of adaptor.tags(node, tag)) {
            if (!adaptor.hasAttribute(element, tag === 'line' ? 'data-line' : 'data-frame')) continue;
            adaptor.setAttribute(element, 'stroke-width', '70');
            adaptor.setAttribute(element, 'fill', 'none');
            const classes = adaptor.getAttribute(element, 'class') || '';
            if (classes.includes('mjx-dashed')) adaptor.setAttribute(element, 'stroke-dasharray', '140');
            if (classes.includes('mjx-dotted')) {
                adaptor.setAttribute(element, 'stroke-linecap', 'round');
                adaptor.setAttribute(element, 'stroke-dasharray', '0,140');
            }
        }
    }
    for (const element of adaptor.tags(node, 'svg').slice(1)) adaptor.setAttribute(element, 'overflow', 'visible');
    const svg = adaptor.innerHTML(node);
    if (/data-mjx-error=/.test(svg)) throw new Error('LaTeX 语法无效，无法生成 SVG。');
    return svg.replace(/stroke="currentColor"/g, 'stroke="#000"')
        .replace(/fill="currentColor"/g, 'fill="#000"').trim();
};
