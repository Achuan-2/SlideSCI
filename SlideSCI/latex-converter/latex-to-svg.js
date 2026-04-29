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
    else if (content.startsWith('$') && content.endsWith('$')) {
        content = content.substring(1, content.length - 1);
    }
    else if (content.startsWith('\\(') && content.endsWith('\\)')) {
        content = content.substring(2, content.length - 2);
    }

    return normalizeCommonLatexTypos(content.trim());
}

function normalizeCommonLatexTypos(latex) {
    let content = latex;

    content = content.replace(/(^|[^\\A-Za-z])xrightarrow\s*([A-Za-z0-9]+)\b/g, '$1\\xrightarrow{$2}');
    content = content.replace(/(^|[^\\A-Za-z])xleftarrow\s*([A-Za-z0-9]+)\b/g, '$1\\xleftarrow{$2}');
    content = content.replace(/\\xrightarrow\s+([A-Za-z0-9]+)\b/g, '\\xrightarrow{$1}');
    content = content.replace(/\\xleftarrow\s+([A-Za-z0-9]+)\b/g, '\\xleftarrow{$1}');
    content = content.replace(/\\xrightarrow([A-Za-z0-9]+)\b/g, '\\xrightarrow{$1}');
    content = content.replace(/\\xleftarrow([A-Za-z0-9]+)\b/g, '\\xleftarrow{$1}');

    return content;
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
        .filter(([name]) => name !== 'style')
        .filter(([name]) => name !== 'role' && name !== 'focusable')
        .filter(([name]) => name === 'data-slidesci-origin' || !name.startsWith('data-'))
        .map(([name, value]) => {
            const normalizedValue = value === 'currentColor' ? '#000' : value;
            return ` ${name}="${escapeAttribute(normalizedValue)}"`;
        })
        .join('');
}

function parseTranslateTransform(transform) {
    if (!transform) {
        return null;
    }

    const match = transform.match(
        /^translate\(\s*([+-]?(?:\d+\.?\d*|\.\d+))(?:[\s,]+([+-]?(?:\d+\.?\d*|\.\d+)))?\s*\)$/
    );
    if (!match) {
        return null;
    }

    const x = Number(match[1]);
    const y = match[2] === undefined ? 0 : Number(match[2]);
    if (!Number.isFinite(x) || !Number.isFinite(y)) {
        return null;
    }

    return { x, y };
}

function identityMatrix() {
    return [1, 0, 0, 1, 0, 0];
}

function multiplyMatrix(left, right) {
    return [
        left[0] * right[0] + left[2] * right[1],
        left[1] * right[0] + left[3] * right[1],
        left[0] * right[2] + left[2] * right[3],
        left[1] * right[2] + left[3] * right[3],
        left[0] * right[4] + left[2] * right[5] + left[4],
        left[1] * right[4] + left[3] * right[5] + left[5],
    ];
}

function transformPoint(matrix, x, y) {
    return {
        x: matrix[0] * x + matrix[2] * y + matrix[4],
        y: matrix[1] * x + matrix[3] * y + matrix[5],
    };
}

function parseTransformMatrix(transform) {
    if (!transform) {
        return identityMatrix();
    }

    let matrix = identityMatrix();
    const pattern = /([a-zA-Z]+)\(([^)]*)\)/g;
    let match;
    let matched = false;

    while ((match = pattern.exec(transform)) !== null) {
        matched = true;
        const name = match[1].toLowerCase();
        const values = match[2].trim().split(/[\s,]+/).filter(Boolean).map(Number);
        if (values.some((value) => !Number.isFinite(value))) {
            return null;
        }

        let next = null;
        if (name === 'translate') {
            next = [1, 0, 0, 1, values[0] || 0, values.length > 1 ? values[1] : 0];
        }
        else if (name === 'scale') {
            const scaleX = values[0];
            const scaleY = values.length > 1 ? values[1] : scaleX;
            next = [scaleX, 0, 0, scaleY, 0, 0];
        }
        else if (name === 'matrix' && values.length === 6) {
            next = values;
        }

        if (!next) {
            return null;
        }

        matrix = multiplyMatrix(matrix, next);
    }

    return matched ? matrix : null;
}

function isIdentityMatrix(matrix) {
    return matrix.every((value, index) => value === identityMatrix()[index]);
}

function serializeMatrixTransform(matrix) {
    return `matrix(${matrix.map(formatNumber).join(' ')})`;
}

function containsCjkText(text) {
    return /[\u3000-\u303F\u3400-\u9FFF\uFF00-\uFFEF]/.test(text || '');
}

function transformPathData(d, matrix) {
    if (!d || isIdentityMatrix(matrix)) {
        return d;
    }

    const tokens = d.match(/[A-Za-z]|[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/g);
    if (!tokens) {
        return d;
    }

    const arity = {
        M: 2, L: 2, T: 2,
        H: 1, V: 1,
        C: 6, S: 4, Q: 4,
        Z: 0,
    };

    const output = [];
    let index = 0;
    let command = null;
    while (index < tokens.length) {
        if (/^[A-Za-z]$/.test(tokens[index])) {
            command = tokens[index++];
            output.push(command);
        }

        if (!command) {
            return d;
        }

        const upper = command.toUpperCase();
        if (command !== upper || !(upper in arity)) {
            return d;
        }

        if (upper === 'Z') {
            command = null;
            continue;
        }

        const count = arity[upper];
        while (index + count <= tokens.length && !/^[A-Za-z]$/.test(tokens[index])) {
            const values = tokens.slice(index, index + count).map(Number);
            if (values.some((value) => !Number.isFinite(value))) {
                return d;
            }

            if (upper === 'H') {
                const point = transformPoint(matrix, values[0], 0);
                output.push(formatNumber(point.x));
            }
            else if (upper === 'V') {
                const point = transformPoint(matrix, 0, values[0]);
                output.push(formatNumber(point.y));
            }
            else {
                for (let i = 0; i < values.length; i += 2) {
                    const point = transformPoint(matrix, values[i], values[i + 1]);
                    output.push(formatNumber(point.x), formatNumber(point.y));
                }
            }

            index += count;

            if (index >= tokens.length || /^[A-Za-z]$/.test(tokens[index])) {
                break;
            }
        }
    }

    return output.join(' ');
}

function parseNumericAttribute(attributes, name, defaultValue) {
    if (attributes[name] === undefined || attributes[name] === null || attributes[name] === '') {
        return defaultValue;
    }

    const value = Number.parseFloat(attributes[name]);
    return Number.isFinite(value) ? value : null;
}

function serializeRectAsPath(attributes, matrix) {
    const x = parseNumericAttribute(attributes, 'x', 0);
    const y = parseNumericAttribute(attributes, 'y', 0);
    const width = parseNumericAttribute(attributes, 'width', null);
    const height = parseNumericAttribute(attributes, 'height', null);
    if (
        x === null ||
        y === null ||
        width === null ||
        height === null ||
        width < 0 ||
        height < 0
    ) {
        return null;
    }

    const topLeft = transformPoint(matrix, x, y);
    const topRight = transformPoint(matrix, x + width, y);
    const bottomRight = transformPoint(matrix, x + width, y + height);
    const bottomLeft = transformPoint(matrix, x, y + height);
    const outputAttributes = { ...attributes };

    delete outputAttributes.x;
    delete outputAttributes.y;
    delete outputAttributes.width;
    delete outputAttributes.height;
    delete outputAttributes.rx;
    delete outputAttributes.ry;
    delete outputAttributes.transform;
    outputAttributes['data-slidesci-origin'] = 'rect';

    const isAxisAligned =
        Math.abs(topLeft.y - topRight.y) < 0.0001 &&
        Math.abs(bottomLeft.y - bottomRight.y) < 0.0001 &&
        Math.abs(topLeft.x - bottomLeft.x) < 0.0001 &&
        Math.abs(topRight.x - bottomRight.x) < 0.0001;

    if (isAxisAligned) {
        const x1 = Math.min(topLeft.x, topRight.x, bottomRight.x, bottomLeft.x);
        const y1 = Math.min(topLeft.y, topRight.y, bottomRight.y, bottomLeft.y);
        const x2 = Math.max(topLeft.x, topRight.x, bottomRight.x, bottomLeft.x);
        const y2 = Math.max(topLeft.y, topRight.y, bottomRight.y, bottomLeft.y);
        outputAttributes.d = `M${formatNumber(x1)} ${formatNumber(y1)}V${formatNumber(y2)}H${formatNumber(x2)}V${formatNumber(y1)}H${formatNumber(x1)}Z`;
    }
    else {
        outputAttributes.d = [
            `M${formatNumber(topLeft.x)} ${formatNumber(topLeft.y)}`,
            `L${formatNumber(topRight.x)} ${formatNumber(topRight.y)}`,
            `L${formatNumber(bottomRight.x)} ${formatNumber(bottomRight.y)}`,
            `L${formatNumber(bottomLeft.x)} ${formatNumber(bottomLeft.y)}`,
            'Z',
        ].join(' ');
    }

    return `<path${serializeAttributes(outputAttributes, new Set(['transform']))}></path>`;
}

function serializeTextElement(node, attributes, currentMatrix) {
    const textContent = adaptor.childNodes(node)
        .map((child) => {
            const childKind = adaptor.kind(child);
            return childKind === '#text' ? escapeText(adaptor.value(child)) : '';
        })
        .join('');
    const outputAttributes = { ...attributes };

    delete outputAttributes.transform;
    const isTranslateOnly =
        Math.abs(currentMatrix[0] - 1) < 0.0001 &&
        Math.abs(currentMatrix[1]) < 0.0001 &&
        Math.abs(currentMatrix[2]) < 0.0001 &&
        Math.abs(currentMatrix[3] - 1) < 0.0001;
    if (isTranslateOnly && !isIdentityMatrix(currentMatrix)) {
        const x = parseNumericAttribute(outputAttributes, 'x', 0);
        const y = parseNumericAttribute(outputAttributes, 'y', 0);
        if (x !== null && y !== null) {
            outputAttributes.x = formatNumber(x + currentMatrix[4]);
            outputAttributes.y = formatNumber(y + currentMatrix[5]);
        }
        else {
            outputAttributes.transform = serializeMatrixTransform(currentMatrix);
        }
    }
    else if (!isIdentityMatrix(currentMatrix)) {
        outputAttributes.transform = serializeMatrixTransform(currentMatrix);
    }

    if (containsCjkText(textContent)) {
        outputAttributes['font-family'] = 'Microsoft YaHei, SimSun, Noto Sans CJK SC, sans-serif';
    }

    return `<text${serializeAttributes(outputAttributes, new Set())}>${textContent}</text>`;
}

function serializeTranslatedBarPath(pathNode, translate) {
    if (!translate) {
        return null;
    }

    const pathAttributes = getAttributes(pathNode);
    const rect = parseRectPath(pathAttributes.d);
    if (!rect) {
        return null;
    }

    const x1 = rect.x + translate.x;
    const y1 = rect.y + translate.y;
    const x2 = x1 + rect.width;
    const y2 = y1 + rect.height;
    const outputAttributes = {
        ...pathAttributes,
        d: `M${formatNumber(x1)} ${formatNumber(y1)}V${formatNumber(y2)}H${formatNumber(x2)}V${formatNumber(y1)}H${formatNumber(x1)}Z`,
    };
    delete outputAttributes.transform;

    return `<path${serializeAttributes(outputAttributes, new Set())}></path>`;
}

function serializeClippedRectPath(pathNode, nestedSvgAttributes, viewBox, matrix) {
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
    let adjustedVisibleX = visibleX;
    let adjustedVisibleRight = visibleRight;
    if (pathAttributes['data-c'] === '2013' && visibleRight - visibleX > 80) {
        const inset = Math.min(170, (visibleRight - visibleX) * 0.24);
        adjustedVisibleX += inset;
        adjustedVisibleRight -= inset;
    }

    const visibleWidth = Math.max(0, adjustedVisibleRight - adjustedVisibleX);
    const visibleHeight = Math.max(0, visibleBottom - visibleY);
    const scaleX = width / viewBox.width;
    const scaleY = height / viewBox.height;
    const pathX = x + (adjustedVisibleX - viewBox.x) * scaleX;
    const pathY = y + (visibleY - viewBox.y) * scaleY;
    const pathRight = pathX + visibleWidth * scaleX;
    const pathBottom = pathY + visibleHeight * scaleY;
    const rawD = `M${formatNumber(pathX)} ${formatNumber(pathY)}V${formatNumber(pathBottom)}H${formatNumber(pathRight)}V${formatNumber(pathY)}H${formatNumber(pathX)}Z`;
    const outputAttributes = {
        ...pathAttributes,
        d: transformPathData(rawD, matrix),
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
    const scaleX = width / viewBox.width;
    const scaleY = height / viewBox.height;
    const translateX = x - viewBox.x * scaleX;
    const translateY = y - viewBox.y * scaleY;
    const nestedMatrix = [
        scaleX,
        0,
        0,
        scaleY,
        translateX,
        translateY,
    ];
    const previousMatrix = state.matrix;
    const mappedMatrix = multiplyMatrix(previousMatrix, nestedMatrix);

    if (children.length === 1 && adaptor.kind(children[0]) === 'path') {
        const clippedPath = serializeClippedRectPath(children[0], attributes, viewBox, previousMatrix);
        if (clippedPath) {
            return clippedPath;
        }
    }

    state.matrix = mappedMatrix;
    const content = children.map((child) => serializeSvgNode(child, state)).join('');
    state.matrix = previousMatrix;

    return content;
}

function applyTranslateToRectPathMarkup(pathMarkup, translate) {
    const dMatch = pathMarkup.match(/ d="([^"]+)"/);
    if (!dMatch) {
        return null;
    }

    const rect = parseRectPath(dMatch[1]);
    if (!rect) {
        return null;
    }

    const x1 = rect.x + translate.x;
    const y1 = rect.y + translate.y;
    const x2 = x1 + rect.width;
    const y2 = y1 + rect.height;
    const newD = `M${formatNumber(x1)} ${formatNumber(y1)}V${formatNumber(y2)}H${formatNumber(x2)}V${formatNumber(y1)}H${formatNumber(x1)}Z`;

    return pathMarkup.replace(/ d="[^"]+"/, ` d="${newD}"`);
}

function serializeElement(node, state, flattenSvg) {
    const kind = adaptor.kind(node);
    const attributes = getAttributes(node);
    const childrenNodes = adaptor.childNodes(node);
    const localMatrix = parseTransformMatrix(attributes.transform) || identityMatrix();
    const previousMatrix = state.matrix;
    const currentMatrix = multiplyMatrix(previousMatrix, localMatrix);
    const renderableChildren = childrenNodes.filter((child) => {
        const childKind = adaptor.kind(child);
        return childKind !== '#text' && childKind !== '#comment';
    });

    if (kind === 'path') {
        const outputAttributes = {
            ...attributes,
            d: transformPathData(attributes.d, currentMatrix),
        };
        delete outputAttributes.transform;
        return `<path${serializeAttributes(outputAttributes, new Set(['transform']))}></path>`;
    }

    if (kind === 'rect') {
        const rectPath = serializeRectAsPath(attributes, currentMatrix);
        if (rectPath) {
            return rectPath;
        }
    }

    if (kind === 'text') {
        return serializeTextElement(node, attributes, currentMatrix);
    }

    state.matrix = currentMatrix;
    const children = childrenNodes.map((child) => serializeSvgNode(child, state)).join('');
    state.matrix = previousMatrix;

    return `<${kind}${serializeAttributes(attributes, new Set(['transform']))}>${children}</${kind}>`;
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
        matrix: identityMatrix(),
        nextClipId: 1,
        svgDepth: 0,
    };

    return adaptor.childNodes(containerNode)
        .map((child) => serializeSvgNode(child, state))
        .join('');
}

function parseRectPathFromMarkup(pathMarkup) {
    const dMatch = pathMarkup.match(/\sd="([^"]+)"/);
    if (!dMatch) {
        return null;
    }

    return parseRectPath(dMatch[1]);
}

function replaceRectPathInMarkup(pathMarkup, rect) {
    const d = `M${formatNumber(rect.x)} ${formatNumber(rect.y)}V${formatNumber(rect.y + rect.height)}H${formatNumber(rect.x + rect.width)}V${formatNumber(rect.y)}H${formatNumber(rect.x)}Z`;
    return pathMarkup.replace(/\sd="[^"]+"/, ` d="${d}"`);
}

function staggerConsecutiveOverlineBars(svgMarkup) {
    const pathPattern = /<path\b[^>]*><\/path>/g;
    const matches = [];
    let match;

    while ((match = pathPattern.exec(svgMarkup)) !== null) {
        const rect = parseRectPathFromMarkup(match[0]);
        if (!rect) {
            continue;
        }
        if (/\sdata-slidesci-origin="rect"/.test(match[0])) {
            continue;
        }
        if (rect.width < 80 || rect.height > 80) {
            continue;
        }

        matches.push({
            index: match.index,
            markup: match[0],
            rect,
        });
    }

    if (matches.length < 3) {
        return svgMarkup;
    }

    const replacements = new Map();
    let group = [];
    const flushGroup = () => {
        if (group.length < 3) {
            group = [];
            return;
        }

        const top = Math.min(...group.map((item) => item.rect.y));
        const thickness = Math.max(...group.map((item) => item.rect.height));
        const step = Math.max(240, thickness * 6.4);

        group.forEach((item, index) => {
            let targetX = item.rect.x;
            let targetWidth = item.rect.width;
            if (index < group.length - 1) {
                const targetEndIndex = index === 0 ? group.length - 1 : index + 1;
                const targetEnd = group[targetEndIndex].rect;
                targetWidth = (targetEnd.x + targetEnd.width) - item.rect.x;
            }

            const targetY = top - step * (group.length - 1 - index);
            replacements.set(item.index, replaceRectPathInMarkup(item.markup, {
                ...item.rect,
                x: targetX,
                width: targetWidth,
                y: targetY,
            }));
        });
        group = [];
    };

    matches.forEach((item) => {
        if (group.length === 0) {
            group.push(item);
            return;
        }

        const previous = group[group.length - 1];
        const gap = item.rect.x - (previous.rect.x + previous.rect.width);
        const averageWidth = (item.rect.width + previous.rect.width) / 2;
        if (gap >= 0 && gap <= averageWidth * 0.9) {
            group.push(item);
        }
        else {
            flushGroup();
            group.push(item);
        }
    });
    flushGroup();

    if (replacements.size === 0) {
        return svgMarkup;
    }

    return svgMarkup.replace(pathPattern, (pathMarkup, offset) => replacements.get(offset) || pathMarkup);
}

function addRootViewBoxPadding(svgMarkup) {
    return svgMarkup.replace(/<svg\b([^>]*)\bviewBox="([^"]+)"([^>]*)>/, (match, before, viewBoxValue, after) => {
        const viewBox = parseViewBox(viewBoxValue);
        if (!viewBox) {
            return match;
        }

        const verticalPadding = Math.max(60, viewBox.height * 0.08);
        const rectBounds = [];
        const pathPattern = /<path\b[^>]*><\/path>/g;
        let pathMatch;
        while ((pathMatch = pathPattern.exec(svgMarkup)) !== null) {
            const rect = parseRectPathFromMarkup(pathMatch[0]);
            if (rect) {
                rectBounds.push(rect);
            }
        }

        const currentLeft = viewBox.x;
        const currentTop = viewBox.y;
        const currentRight = viewBox.x + viewBox.width;
        const currentBottom = viewBox.y + viewBox.height;
        const rectLeft = rectBounds.length > 0 ? Math.min(...rectBounds.map((rect) => rect.x)) : currentLeft;
        const rectTop = rectBounds.length > 0 ? Math.min(...rectBounds.map((rect) => rect.y)) : currentTop;
        const rectRight = rectBounds.length > 0
            ? Math.max(...rectBounds.map((rect) => rect.x + rect.width))
            : currentRight;
        const rectBottom = rectBounds.length > 0
            ? Math.max(...rectBounds.map((rect) => rect.y + rect.height))
            : currentBottom;

        const paddedLeft = Math.min(currentLeft, rectLeft);
        const paddedTop = Math.min(currentTop - verticalPadding, rectTop - verticalPadding);
        const paddedRight = Math.max(currentRight, rectRight);
        const paddedBottom = Math.max(currentBottom + verticalPadding, rectBottom + verticalPadding);
        const paddedViewBox = [
            formatNumber(paddedLeft),
            formatNumber(paddedTop),
            formatNumber(paddedRight - paddedLeft),
            formatNumber(paddedBottom - paddedTop),
        ].join(' ');

        return `<svg${before} viewBox="${paddedViewBox}"${after} overflow="visible">`;
    });
}

function convertLatexToSvg(latex, displayMode) {
    const node = html.convert(latex, {
        display: displayMode,
        em: 20,
        ex: 10,
        containerWidth: 80 * 20,
    });

    // Office/WPS can misplace MathJax's nested SVG nodes that are used for
    // bars, braces and extensible arrows. Flatten them into one SVG coordinate
    // system and add a small viewBox margin to avoid clipping accents.
    const processedSvg = addRootViewBoxPadding(staggerConsecutiveOverlineBars(serializeOfficeCompatibleSvg(node)))
        .replace(/\sdata-slidesci-origin="rect"/g, '')
        .replace(/stroke="currentColor"/g, 'stroke="#000"')
        .replace(/fill="currentColor"/g, 'fill="#000"');

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
