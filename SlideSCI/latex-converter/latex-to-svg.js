// CLI mode remains available; --server reuses MathJax modules over JSON lines.
let runtime;
let SvgRenderer;

function loadRuntime() {
    if (runtime) return runtime;
    try {
        const { mathjax } = require('mathjax-full/js/mathjax.js');
        const { TeX } = require('mathjax-full/js/input/tex.js');
        const { AllPackages } = require('mathjax-full/js/input/tex/AllPackages.js');
        const { liteAdaptor } = require('mathjax-full/js/adaptors/liteAdaptor.js');
        const { RegisterHTMLHandler } = require('mathjax-full/js/handlers/html.js');
        const { STATE } = require('mathjax-full/js/core/MathItem.js');
        const { SerializedMmlVisitor } = require('mathjax-full/js/core/MmlTree/SerializedMmlVisitor.js');
        const adaptor = liteAdaptor();
        RegisterHTMLHandler(adaptor);
        runtime = { mathjax, TeX, AllPackages, adaptor, STATE, SerializedMmlVisitor };
        return runtime;
    } catch (error) {
        if (error.code === 'MODULE_NOT_FOUND') {
            throw new Error(`缺少 MathJax 依赖，请在 ${__dirname} 中运行 pnpm install。\n${error.message}`);
        }
        throw error;
    }
}

function createSvgRenderer() {
    if (!SvgRenderer) SvgRenderer = require('mathjax-full/js/output/svg.js').SVG;
    return new SvgRenderer({ fontCache: 'none' });
}

function createDocument(mathMl) {
    const { mathjax, TeX, AllPackages } = loadRuntime();
    // Fresh parsers prevent user-defined macros leaking into later insertions.
    const options = { InputJax: new TeX({ packages: AllPackages }) };
    if (!mathMl) {
        options.OutputJax = createSvgRenderer();
    }
    const document = mathjax.document('', options);
    if (mathMl) {
        // bussproofs requires this API even for ordinary formulas, but only proof trees use metrics.
        let metrics;
        document.outputJax.getBBox = (math, doc) => {
            if (!metrics) {
                metrics = createSvgRenderer();
                metrics.setAdaptor(loadRuntime().adaptor);
                metrics.initialize();
            }
            return metrics.getBBox(math, doc);
        };
    }
    return document;
}

function shouldUseDisplayMode(latex) {
    const trimmed = latex.trim();
    return trimmed.startsWith('\\[') || trimmed.startsWith('$$')
        || /\\begin\s*\{(align|equation|gather|multline|cases|matrix)/.test(trimmed)
        || /\\\\/.test(trimmed);
}

function normalizeLatex(latex) {
    let content = latex.trim();
    if ((content.startsWith('$$') && content.endsWith('$$'))
        || (content.startsWith('\\[') && content.endsWith('\\]'))
        || (content.startsWith('\\(') && content.endsWith('\\)'))) {
        content = content.substring(2, content.length - 2);
    } else if (content.startsWith('$') && content.endsWith('$')) {
        content = content.substring(1, content.length - 1);
    }
    return content.trim();
}

function convertSources(sources, mathMl, explicitDisplay) {
    if (!Array.isArray(sources) || sources.length === 0) {
        throw new Error('公式列表不能为空。');
    }
    const document = createDocument(mathMl);
    return sources.map((source) => {
        if (typeof source !== 'string' || !normalizeLatex(source)) {
            throw new Error('LaTeX 输入为空。');
        }
        const latex = normalizeLatex(source);
        const display = explicitDisplay !== undefined ? explicitDisplay : shouldUseDisplayMode(source);
        if (mathMl) {
            const { STATE, SerializedMmlVisitor } = loadRuntime();
            const root = document.convert(latex, { display, end: STATE.COMPILED });
            const result = new SerializedMmlVisitor().visitTree(root);
            if (/<merror(?:\s|>)/.test(result)) {
                throw new Error('LaTeX 语法无效，无法生成可编辑公式。');
            }
            return result;
        }
        const node = document.convert(latex, { display, em: 20, ex: 10, containerWidth: 1600 });
        return loadRuntime().adaptor.innerHTML(node)
            .replace(/stroke="currentColor"/g, 'stroke="#000"')
            .replace(/fill="currentColor"/g, 'fill="#000"')
            .replace(/\n{2,}/g, '\n').trim();
    });
}

function writeResponse(response) {
    process.stdout.write(JSON.stringify(response) + '\n');
}

function runServer() {
    // Warm only MathML. SVG fonts are loaded when SVG is actually requested.
    convertSources(['x'], true, false);
    const input = require('readline').createInterface({ input: process.stdin, crlfDelay: Infinity });
    input.on('line', (line) => {
        try {
            const request = JSON.parse(line.replace(/^\uFEFF/, ''));
            const results = convertSources(request.sources, request.mathMl === true, request.display);
            writeResponse({ ok: true, results });
        } catch (error) {
            // Invalid input must not terminate a worker used by later insertions.
            writeResponse({ ok: false, error: error.message });
        }
    });
    input.on('error', (error) => {
        console.error(error.message);
        process.exitCode = 1;
        input.close();
    });
    writeResponse({ ready: true });
}

function parseArgs(argv) {
    const options = { display: undefined, mathMl: false, json: false, server: false, latex: null };
    const source = [];
    for (const arg of argv.slice(2)) {
        if (arg === '--server') options.server = true;
        else if (arg === '--mathml') options.mathMl = true;
        else if (arg === '--json') options.json = true;
        else if (arg === '--display') options.display = true;
        else if (arg === '--inline') options.display = false;
        else source.push(arg);
    }
    options.latex = source.length > 0 ? source.join(' ') : null;
    return options;
}

function runCli(source, options) {
    const sources = options.json ? JSON.parse(source.replace(/^\uFEFF/, '')) : [source];
    const results = convertSources(sources, options.mathMl, options.display);
    process.stdout.write(options.json ? JSON.stringify(results) : results[0]);
}

function reportError(error) {
    console.error(`LaTeX 转换失败: ${error.message}`);
    process.exitCode = 1;
}

try {
    const options = parseArgs(process.argv);
    if (options.server) {
        runServer();
    } else if (!process.stdin.isTTY) {
        let input = '';
        process.stdin.setEncoding('utf8');
        process.stdin.on('data', (chunk) => { input += chunk; });
        process.stdin.on('end', () => {
            try { runCli(input || options.latex || '', options); } catch (error) { reportError(error); }
        });
        process.stdin.on('error', reportError);
    } else {
        runCli(options.latex || '', options);
    }
} catch (error) {
    reportError(error);
}
