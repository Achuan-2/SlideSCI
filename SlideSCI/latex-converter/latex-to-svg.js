// 检查依赖是否存在
try {
    const { mathjax } = require('mathjax-full/js/mathjax.js');
    const { TeX } = require('mathjax-full/js/input/tex.js');
    const { SVG } = require('mathjax-full/js/output/svg.js');
    const { liteAdaptor } = require('mathjax-full/js/adaptors/liteAdaptor.js');
    const { RegisterHTMLHandler } = require('mathjax-full/js/handlers/html.js');
    const { AllPackages } = require('mathjax-full/js/input/tex/AllPackages.js');
} catch (error) {
    console.error('缺少必要的依赖包 mathjax-full。请按以下步骤安装：');
    console.error('1. 确认已安装 Node.js（下载地址：https://nodejs.org/）');
    console.error('2. 打开命令提示符，导航到当前目录：' + __dirname);
    console.error('3. 然后运行：npm install');
    console.error('4. 安装完成后重试LaTeX转换功能');
    console.error('');
    process.exit(1);
}

const { mathjax } = require('mathjax-full/js/mathjax.js');
const { TeX } = require('mathjax-full/js/input/tex.js');
const { SVG } = require('mathjax-full/js/output/svg.js');
const { liteAdaptor } = require('mathjax-full/js/adaptors/liteAdaptor.js');
const { RegisterHTMLHandler } = require('mathjax-full/js/handlers/html.js');
const { AllPackages } = require('mathjax-full/js/input/tex/AllPackages.js');

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);

const tex = new TeX({
    packages: AllPackages,
    inlineMath: [
        ['\\(', '\\)'],
        ['$', '$'],
    ],
    displayMath: [
        ['\\[', '\\]'],
        ['$$', '$$'],
    ],

});

const svg = new SVG({
    fontCache: 'none',
});

const html = mathjax.document('', {
    InputJax: tex,
    OutputJax: svg,
});

function shouldUseDisplayMode(latex) {
    const trimmed = latex.trim();
    if (trimmed.startsWith('\\[') || trimmed.startsWith('$$')) {
        return true;
    }

    return /\\begin\s*\{(align|equation|gather|multline|cases|matrix)/.test(trimmed) || /\\\\/.test(trimmed);
}

function normalizeLatex(latex) {
    let content = latex.trim();

    if (
        (content.startsWith('$$') && content.endsWith('$$')) ||
        (content.startsWith('\\[') && content.endsWith('\\]'))
    ) {
        content = content.substring(2, content.length - 2);
    }
    else if (
        (content.startsWith('$') && content.endsWith('$')) ||
        (content.startsWith('\\(') && content.endsWith('\\)'))
    ) {
        content = content.substring(1, content.length - 1);
    }

    return content.trim();
}

function escapeAttribute(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function escapeText(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function formatNumber(value) {
    if (!Number.isFinite(value)) {
        return '0';
    }

    return Number(value.toFixed(6)).toString();
}

function getAttributes(node) {
    const attributes = {};
    adaptor.allAttributes(node).forEach((attr) => {
        attributes[attr.name] = attr.value;
    });
    return attributes;
}

function parseViewBox(viewBox) {
    if (!viewBox) {
        return null;
    }

    const values = viewBox.trim().split(/[\s,]+/).map(Number);
    if (values.length !== 4 || values.some((value) => !Number.isFinite(value))) {
        return null;
    }

    return {
        x: values[0],
        y: values[1],
        width: values[2],
        height: values[3],
    };
}

function parseScaleTransform(transform) {
    if (!transform) {
        return { x: 1, y: 1 };
    }

    const match = transform.match(/^scale\(\s*([+-]?(?:\d+\.?\d*|\.\d+))(?:[\s,]+([+-]?(?:\d+\.?\d*|\.\d+)))?\s*\)$/);
    if (!match) {
        return null;
    }

    const scaleX = Number(match[1]);
    const scaleY = match[2] === undefined ? scaleX : Number(match[2]);
    if (!Number.isFinite(scaleX) || !Number.isFinite(scaleY)) {
        return null;
    }

    return { x: scaleX, y: scaleY };
}

function parseRectPath(d) {
    if (!d) {
        return null;
    }

    const match = d.match(/^M\s*([+-]?(?:\d+\.?\d*|\.\d+))\s+([+-]?(?:\d+\.?\d*|\.\d+))\s*V\s*([+-]?(?:\d+\.?\d*|\.\d+))\s*H\s*([+-]?(?:\d+\.?\d*|\.\d+))\s*V\s*([+-]?(?:\d+\.?\d*|\.\d+))\s*H\s*([+-]?(?:\d+\.?\d*|\.\d+))\s*Z$/);
    if (!match) {
        return null;
    }

    const x1 = Number(match[1]);
    const y1 = Number(match[2]);
    const y2 = Number(match[3]);
    const x2 = Number(match[4]);
    const y3 = Number(match[5]);
    const x3 = Number(match[6]);
    if ([x1, y1, y2, x2, y3, x3].some((value) => !Number.isFinite(value))) {
        return null;
    }

    return {
        x: Math.min(x1, x2, x3),
        y: Math.min(y1, y2, y3),
        width: Math.max(x1, x2, x3) - Math.min(x1, x2, x3),
        height: Math.max(y1, y2, y3) - Math.min(y1, y2, y3),
    };
}

function serializeAttributes(attributes, excludedNames) {
    return Object.entries(attributes)
        .filter(([name]) => !excludedNames.has(name))
        .map(([name, value]) => {
            const normalizedValue = value === 'currentColor' ? '#000' : value;
            return ` ${name}="${escapeAttribute(normalizedValue)}"`;
        })
        .join('');
}

function serializeClippedRectPath(pathNode, nestedSvgAttributes, viewBox) {
    const width = Number.parseFloat(nestedSvgAttributes.width);
    const height = Number.parseFloat(nestedSvgAttributes.height);
    const x = Number.parseFloat(nestedSvgAttributes.x || '0');
    const y = Number.parseFloat(nestedSvgAttributes.y || '0');
    if (![width, height, x, y].every(Number.isFinite) || viewBox.width === 0 || viewBox.height === 0) {
        return null;
    }

    const pathAttributes = getAttributes(pathNode);
    const rect = parseRectPath(pathAttributes.d);
    const scale = parseScaleTransform(pathAttributes.transform);
    if (!rect || !scale) {
        return null;
    }

    const scaledRect = {
        x: rect.x * scale.x,
        y: rect.y * scale.y,
        width: rect.width * Math.abs(scale.x),
        height: rect.height * Math.abs(scale.y),
    };
    const visibleX = Math.max(scaledRect.x, viewBox.x);
    const visibleY = Math.max(scaledRect.y, viewBox.y);
    const visibleRight = Math.min(scaledRect.x + scaledRect.width, viewBox.x + viewBox.width);
    const visibleBottom = Math.min(scaledRect.y + scaledRect.height, viewBox.y + viewBox.height);
    const visibleWidth = Math.max(0, visibleRight - visibleX);
    const visibleHeight = Math.max(0, visibleBottom - visibleY);
    const scaleX = width / viewBox.width;
    const scaleY = height / viewBox.height;
    const pathX = x + (visibleX - viewBox.x) * scaleX;
    const pathY = y + (visibleY - viewBox.y) * scaleY;
    const pathRight = pathX + visibleWidth * scaleX;
    const pathBottom = pathY + visibleHeight * scaleY;
    const outputAttributes = {
        ...pathAttributes,
        d: `M${formatNumber(pathX)} ${formatNumber(pathY)}V${formatNumber(pathBottom)}H${formatNumber(pathRight)}V${formatNumber(pathY)}H${formatNumber(pathX)}Z`,
    };
    delete outputAttributes.transform;

    return `<path${serializeAttributes(outputAttributes, new Set())}></path>`;
}

function serializeNestedSvg(node, state) {
    const attributes = getAttributes(node);
    const viewBox = parseViewBox(attributes.viewBox);
    const width = Number.parseFloat(attributes.width);
    const height = Number.parseFloat(attributes.height);
    const x = Number.parseFloat(attributes.x || '0');
    const y = Number.parseFloat(attributes.y || '0');
    if (!viewBox || ![width, height, x, y].every(Number.isFinite) || viewBox.width === 0 || viewBox.height === 0) {
        return serializeElement(node, state, false);
    }

    const children = adaptor.childNodes(node);
    if (children.length === 1 && adaptor.kind(children[0]) === 'path') {
        const clippedPath = serializeClippedRectPath(children[0], attributes, viewBox);
        if (clippedPath) {
            return clippedPath;
        }
    }

    const clipId = `slidesci-clip-${state.nextClipId++}`;
    const scaleX = width / viewBox.width;
    const scaleY = height / viewBox.height;
    const translateX = x - viewBox.x * scaleX;
    const translateY = y - viewBox.y * scaleY;
    const matrix = [
        formatNumber(scaleX),
        '0',
        '0',
        formatNumber(scaleY),
        formatNumber(translateX),
        formatNumber(translateY),
    ].join(' ');
    const content = children.map((child) => serializeSvgNode(child, state)).join('');

    return [
        `<defs><clipPath id="${clipId}" clipPathUnits="userSpaceOnUse">`,
        `<rect x="${formatNumber(viewBox.x)}" y="${formatNumber(viewBox.y)}" width="${formatNumber(viewBox.width)}" height="${formatNumber(viewBox.height)}"></rect>`,
        '</clipPath></defs>',
        `<g transform="matrix(${matrix})"><g clip-path="url(#${clipId})">${content}</g></g>`,
    ].join('');
}

function serializeElement(node, state, flattenSvg) {
    const kind = adaptor.kind(node);
    const attributes = getAttributes(node);
    const children = adaptor.childNodes(node).map((child) => serializeSvgNode(child, state)).join('');

    return `<${kind}${serializeAttributes(attributes, new Set())}>${children}</${kind}>`;
}

function serializeSvgNode(node, state) {
    const kind = adaptor.kind(node);
    if (kind === '#text') {
        return escapeText(adaptor.value(node));
    }

    if (kind === '#comment') {
        return '';
    }

    if (kind === 'svg') {
        if (state.svgDepth > 0) {
            return serializeNestedSvg(node, state);
        }

        state.svgDepth += 1;
        const output = serializeElement(node, state, false);
        state.svgDepth -= 1;
        return output;
    }

    return serializeElement(node, state, false);
}

function serializeOfficeCompatibleSvg(containerNode) {
    const state = {
        nextClipId: 1,
        svgDepth: 0,
    };

    return adaptor.childNodes(containerNode)
        .map((child) => serializeSvgNode(child, state))
        .join('');
}

function convertLatexToSvg(latex, displayMode) {
    const node = html.convert(latex, {
        display: displayMode,
        em: 20,
        ex: 10,
        containerWidth: 80 * 20,
    });

    const processedSvg = serializeOfficeCompatibleSvg(node);

    return processedSvg.replace(/\n{2,}/g, '\n').trim();
}

function runConversion(rawLatex, explicitDisplay) {
    const normalized = normalizeLatex(rawLatex);
    const displayMode = explicitDisplay !== undefined ? explicitDisplay : shouldUseDisplayMode(rawLatex);

    if (!normalized) {
        console.error('LaTeX 输入为空。');
        process.exit(1);
    }

    try {
        const svgOutput = convertLatexToSvg(normalized, displayMode);
        process.stdout.write(svgOutput);
    }
    catch (error) {
        console.error(`LaTeX 转换失败: ${error.message}`);
        console.error('');
        console.error('请检查：');
        console.error('1. LaTeX 语法是否正确');
        console.error('2. 是否包含不支持的命令或包');
        console.error('3. 如果是首次使用，请确认已正确安装依赖（运行 npm install）');
        process.exit(1);
    }
}

function parseArgs(argv) {
    const args = argv.slice(2);
    const options = {
        display: undefined,
        latex: null,
    };

    args.forEach((arg) => {
        if (arg === '--display') {
            options.display = true;
        }
        else if (arg === '--inline') {
            options.display = false;
        }
        else if (!options.latex) {
            options.latex = arg;
        }
        else {
            options.latex += ` ${arg}`;
        }
    });

    return options;
}

(function main() {
    const options = parseArgs(process.argv);

    if (!process.stdin.isTTY) {
        let input = '';
        process.stdin.setEncoding('utf8');
        process.stdin.on('data', (chunk) => {
            input += chunk;
        });
        process.stdin.on('end', () => {
            const source = input.length > 0 ? input : options.latex;
            if (!source) {
                console.error('未提供任何 LaTeX 输入。');
                process.exit(1);
            }
            runConversion(source, options.display);
        });
        process.stdin.on('error', (error) => {
            console.error(`读取标准输入失败: ${error.message}`);
            console.error('请确认：');
            console.error('1. 输入的LaTeX内容格式正确');
            console.error('2. Node.js环境正常运行');
            process.exit(1);
        });
    }
    else if (options.latex) {
        runConversion(options.latex, options.display);
    }
    else {
        console.error('未提供任何 LaTeX 输入。');
        process.exit(1);
    }
})();
