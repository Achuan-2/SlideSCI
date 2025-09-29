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

function convertLatexToSvg(latex, displayMode) {
    const node = html.convert(latex, {
        display: displayMode,
        em: 20,
        ex: 10,
        containerWidth: 80 * 20,
    });

    const svgContent = adaptor.innerHTML(node);
    return svgContent.replace(/\n{2,}/g, '\n').trim();
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
