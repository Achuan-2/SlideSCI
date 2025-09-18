using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq; // Add this line
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    public class CodeHighlighter
    {
        private Dictionary<string, Color> themeColors;
        private Dictionary<
            string,
            List<(string pattern, RegexOptions options, string type)>
        > languagePatterns;
        private Dictionary<string, string> languageAliases;
        private HashSet<string> processedRanges;

        public CodeHighlighter(bool isDarkTheme)
        {
            processedRanges = new HashSet<string>();
            InitializeColors(isDarkTheme);
            InitializePatterns();
            InitializeAliases();
        }

        private void InitializeColors(bool isDarkTheme)
        {
            if (isDarkTheme)
            {
                themeColors = new Dictionary<string, Color>
                {
                    { "keyword", Color.FromArgb(86, 156, 214) }, // 蓝色
                    { "comment", Color.FromArgb(87, 166, 74) }, // 绿色
                    { "string", Color.FromArgb(214, 157, 133) }, // 棕色
                    { "number", Color.FromArgb(181, 206, 168) }, // 蓝绿色
                    { "property", Color.FromArgb(156, 220, 254) }, // 浅蓝色
                    { "selector", Color.FromArgb(215, 186, 125) }, // 金色
                };
            }
            else
            {
                themeColors = new Dictionary<string, Color>
                {
                    { "keyword", Color.FromArgb(0, 0, 255) }, // 蓝色
                    { "comment", Color.FromArgb(0, 128, 0) }, // 绿色
                    { "string", Color.FromArgb(163, 21, 21) }, // 棕色
                    { "number", Color.FromArgb(9, 134, 88) }, // 蓝绿色
                    { "property", Color.FromArgb(0, 134, 209) }, // 浅蓝色
                    { "selector", Color.FromArgb(168, 99, 0) }, // 褐色
                };
            }
        }

        private void InitializePatterns()
        {
            languagePatterns = new Dictionary<
                string,
                List<(string pattern, RegexOptions options, string type)>
            >
            {
                {
                    "csharp",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"/\*[\s\S]*?\*/", RegexOptions.None, "comment"),
                        (@"//[^\n]*", RegexOptions.None, "comment"),
                        // 字符串
                        (@"@""([^""]|"""")*""|""([^""\n\\]|\\.)*""", RegexOptions.None, "string"),
                        // 数字
                        (@"\b\d*\.?\d+([eE][-+]?\d+)?\b", RegexOptions.None, "number"),
                        // 关键字
                        (
                            @"\b(abstract|as|base|bool|break|byte|case|catch|char|checked|class|const|continue|decimal|default|delegate|do|double|else|enum|event|explicit|extern|false|finally|fixed|float|for|foreach|goto|if|implicit|in|int|interface|internal|is|lock|long|namespace|new|null|object|operator|out|override|params|private|protected|public|readonly|ref|return|sbyte|sealed|short|sizeof|stackalloc|static|string|struct|switch|this|throw|true|try|typeof|uint|ulong|unchecked|unsafe|ushort|using|virtual|void|volatile|while)\b",
                            RegexOptions.None,
                            "keyword"
                        ),
                    }
                },
                {
                    "python",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"#.*?$", RegexOptions.Multiline, "comment"),
                        // 字符串
                        (
                            @"(?:[ruf]|rf|fr)?(?:'''[\s\S]*?'''|\""""""[\s\S]*?\""""""|\x27[^\x27\n\\]*(?:\\.[^\x27\n\\]*)*\x27|\""[^\""\n\\]*(?:\\.[^\""\n\\]*)*\"")",
                            RegexOptions.None,
                            "string"
                        ),
                        // 数字
                        (@"\b\d*\.?\d+([eE][-+]?\d+)?\b", RegexOptions.None, "number"),
                        (
                            @"\b(and|as|assert|async|await|break|class|continue|def|del|elif|else|except|False|finally|for|from|global|if|import|in|is|lambda|None|nonlocal|not|or|pass|raise|return|True|try|while|with|yield)\b",
                            RegexOptions.None,
                            "keyword"
                        ),
                    }
                },
                {
                    "javascript",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"/\*[\s\S]*?\*/|//.*?$", RegexOptions.Multiline, "comment"),
                        // 字符串
                        (
                            @"`(?:[^`\\]|\\.)*`|'(?:[^'\\]|\\.)*'|""(?:[^""\\]|\\.)*""",
                            RegexOptions.None,
                            "string"
                        ),
                        // 数字
                        (@"\b\d*\.?\d+([eE][-+]?\d+)?\b", RegexOptions.None, "number"),
                        (
                            @"\b(async|await|break|case|catch|class|const|continue|debugger|default|delete|do|else|export|extends|finally|for|function|if|import|in|instanceof|new|return|super|switch|this|throw|try|typeof|var|void|while|with|yield|let)\b",
                            RegexOptions.None,
                            "keyword"
                        ),
                    }
                },
                {
                    "matlab",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"%.*?$", RegexOptions.Multiline, "comment"),
                        // 字符串 (支持单引号和双引号)
                        (@"(?:""[^""\n]*""|'[^'\n]*')", RegexOptions.None, "string"),
                        // 数字
                        (@"\b\d*\.?\d+([eE][-+]?\d+)?\b", RegexOptions.None, "number"),
                        // 关键字
                        (
                            @"\b(break|case|catch|classdef|continue|else|elseif|end|for|function|global|if|otherwise|parfor|persistent|return|switch|try|while|clear|close|load|save|figure|plot|xlabel|ylabel|title|grid|hold|zeros|ones|rand|eye|disp|input|fprintf|strcmp|length|size|max|min|sum|mean|std|find|sort|reshape)\b",
                            RegexOptions.None,
                            "keyword"
                        ),
                    }
                },
                {
                    "css",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"/\*[\s\S]*?\*/", RegexOptions.None, "comment"),
                        // CSS选择器 - 在大括号前的任何非大括号内容
                        (@"[^{}]+(?=\s*{)", RegexOptions.None, "selector"),
                        // CSS属性
                        (@"[\w-]+(?=\s*:)", RegexOptions.None, "property"),
                        // 字符串
                        (@"'[^']*'|""[^""]*""", RegexOptions.None, "string"),
                        // 数字和单位
                        (
                            @"\b\d*\.?\d+(?:px|em|rem|vh|vw|%|s|ms|deg|rad|turn)?\b",
                            RegexOptions.None,
                            "number"
                        ),
                        // 关键字
                        (
                            @"\b(important|inherit|initial|unset|none|auto|hidden|visible|block|inline|flex|grid|absolute|relative|fixed|static|left|right|top|bottom|center|justify|stretch|wrap|nowrap|solid|dashed|dotted)\b",
                            RegexOptions.None,
                            "keyword"
                        ),
                    }
                },
                {
                    "html",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"<!--[\s\S]*?-->", RegexOptions.None, "comment"),
                        // 字符串
                        (@"'[^']*'|""[^""]*""", RegexOptions.None, "string"),
                        // HTML标签
                        (
                            @"</?\w+(?:\s+\w+(?:\s*=\s*(?:"".*?""|'.*?'|[\^'""\s]\S+))?)*\s*/?>",
                            RegexOptions.None,
                            "keyword"
                        ),
                        // 数字
                        (@"\b\d+\b", RegexOptions.None, "number"),
                    }
                },
                {
                    "r",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"#.*?$", RegexOptions.Multiline, "comment"),
                        // 字符串 (支持单引号和双引号)
                        (
                            @"(?:""[^""\n\\]*(?:\\.[^""\n\\]*)*""|'[^'\n\\]*(?:\\.[^'\n\\]*)*')",
                            RegexOptions.None,
                            "string"
                        ),
                        // 数字
                        (@"\b\d*\.?\d+([eE][-+]?\d+)?L?\b", RegexOptions.None, "number"),
                        // 关键字和函数
                        (
                            @"\b(if|else|for|in|while|function|repeat|next|break|TRUE|FALSE|NULL|Inf|NaN|NA|NA_integer_|NA_real_|NA_complex_|NA_character_|c|list|data\.frame|matrix|array|factor|length|names|dim|class|str|summary|head|tail|print|cat|paste|paste0|substr|nchar|grep|gsub|which|is\.na|is\.null|is\.numeric|is\.character|is\.logical|as\.numeric|as\.character|as\.logical|mean|median|sd|var|min|max|sum|apply|lapply|sapply|mapply|tapply|aggregate|merge|rbind|cbind|subset|sort|order|unique|duplicated|table|plot|hist|boxplot|barplot|pie|lines|points|abline|legend|title|xlabel|ylabel|par|dev\.new|dev\.off|png|pdf|jpeg|library|require|install\.packages|source|load|save|write\.csv|read\.csv|read\.table|write\.table)\b",
                            RegexOptions.None,
                            "keyword"
                        ),
                    }
                },
                {
                    "fortran",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 最高优先级，整行匹配（支持C风格注释和!注释）
                        (@"!.*?$", RegexOptions.Multiline | RegexOptions.IgnoreCase, "comment"),
                        // 字符串 - 只在非注释行中匹配
                        (
                            @"(?:""[^""\n\\]*(?:\\.[^""\n\\]*)*""|'[^'\n\\]*(?:\\.[^'\n\\]*)*')",
                            RegexOptions.None,
                            "string"
                        ),
                        // 数字 (包括科学计数法和Fortran特有的D记号)
                        (@"\b\d*\.?\d+([eEdD][-+]?\d+)?\b", RegexOptions.None, "number"),
                        // Fortran特殊短语关键字 (必须完整匹配)
                        (
                            @"\b(?i)implicit\s+none\b",
                            RegexOptions.IgnoreCase,
                            "keyword"
                        ),
                        (
                            @"\b(?i)double\s+precision\b",
                            RegexOptions.IgnoreCase,
                            "keyword"
                        ),
                        (
                            @"\b(?i)end\s+(program|subroutine|function|module|type|interface|do|if|select)\b",
                            RegexOptions.IgnoreCase,
                            "keyword"
                        ),
                        // Fortran关键字 (包括现代Fortran和传统Fortran)
                        (
                            @"\b(?i)(program|end|subroutine|function|module|use|implicit|none|integer|real|complex|logical|character|parameter|dimension|allocatable|pointer|target|intent|in|out|inout|optional|public|private|save|data|common|equivalence|external|intrinsic|interface|contains|procedure|abstract|extends|class|type|select|case|default|where|elsewhere|forall|pure|elemental|recursive|result|only|operator|assignment|generic|sequence|bind|value|volatile|asynchronous|protected|enum|enumerator|associate|block|critical|sync|all|images|memory|lock|unlock|event|post|wait|co_broadcast|co_max|co_min|co_sum|co_reduce|atomic|define|ref|cas|fetch|add|and|or|xor|if|then|else|elseif|endif|do|while|enddo|continue|exit|cycle|stop|pause|return|call|goto|assign|to|format|open|close|read|write|print|rewind|backspace|endfile|inquire|namelist|include|import)\b",
                            RegexOptions.IgnoreCase,
                            "keyword"
                        ),
                        // 内置函数和过程
                        (
                            @"\b(?i)(abs|acos|aimag|aint|alog|alog10|amax0|amax1|amin0|amin1|amod|anint|asin|atan|atan2|cabs|ccos|cexp|char|clog|cmplx|conjg|cos|cosh|csin|csqrt|dabs|dacos|dasin|datan|datan2|dble|dcos|dcosh|ddim|dexp|dim|dint|dlog|dlog10|dmax1|dmin1|dmod|dnint|dprod|dsign|dsin|dsinh|dsqrt|dtan|dtanh|exp|float|iabs|ichar|idim|idint|idnint|ifix|index|int|isign|len|lge|lgt|lle|llt|log|log10|max|max0|max1|min|min0|min1|mod|nint|real|sign|sin|sinh|sngl|sqrt|tan|tanh|trim|adjustl|adjustr|all|any|count|maxval|minval|product|sum|matmul|dot_product|pack|unpack|reshape|spread|merge|eoshift|cshift|transpose|lbound|ubound|size|shape|allocated|associated|present|kind|selected_int_kind|selected_real_kind|huge|tiny|epsilon|precision|radix|range|digits|minexponent|maxexponent|fraction|exponent|spacing|rrspacing|scale|set_exponent|nearest|ceiling|floor|modulo|sign|verify|scan|null|transfer|bit_size|btest|iand|ibclr|ibits|ibset|ieor|ior|ishft|ishftc|not|mvbits|random_number|random_seed|system_clock|date_and_time|cpu_time)\b",
                            RegexOptions.IgnoreCase,
                            "property"
                        ),
                    }
                },
                    // ===== Stata (do/ado) =====
                {
                    "stata",
                    new List<(string, RegexOptions, string)>
                    {
                        // Comments (order matters: block, //, then leading-asterisk)
                        (@"\/\*[\s\S]*?\*\/", RegexOptions.None, "comment"),
                        (@"//.*?$", RegexOptions.Multiline, "comment"),
                        (@"^\s*\*.*?$", RegexOptions.Multiline, "comment"),

                        // Strings (Stata uses double quotes; allow doubled "" and simple single quotes frequently used in text/macros)
                        (@"(?:""([^""\n]|"""")*""|'([^'\n\\]|\\.)*')", RegexOptions.None, "string"),

                        // Numbers (incl. scientific)
                        (@"\b\d*\.?\d+([eE][-+]?\d+)?\b", RegexOptions.None, "number"),

                        // Macros & globals: `local' and $global
                        (@"`[A-Za-z_]\w*'", RegexOptions.None, "property"),
                        (@"\$\w+", RegexOptions.None, "property"),

                        // Built-in functions like abs(), ln(), exp(), real(), string(), regexm(), inlist(), cond(), etc.
                        (@"\b(abs|acos|asin|atan|atan2|ceil|floor|round|ln|log|exp|sqrt|int|real|string|substr|length|lower|upper|proper|trim|itrim|ltrim|rtrim|regexm|regexr|regexs|inlist|cond|min|max|sum|mean|sd|variance|missing|runiform|rnormal)\s*(?=\()", RegexOptions.IgnoreCase, "property"),

                        // Common commands / keywords (case-insensitive)
                        (@"\b(?i)(use|clear|set|version|cd|pwd|dir|which|do|run|quietly|noi?sy|capture|program|end|syntax|args|local|global|tempvar|tempname|tempfile|macro|display|di|if|in|by|bysort|sort|gsort|keep|drop|rename|order|label|label\s+define|label\s+values|encode|decode|tostring|destring|egen|generate|gen|replace|recode|collapse|contract|expand|append|merge|m:?1|m:?n|joinby|cross|reshape|reshape\s+wide|reshape\s+long|frame|frame\s+create|frame\s+change|frame\s+put|frame\s+drop|preserve|restore|save|saveold|use|import|import\s+excel|export|export\s+excel|insheet|outsheet|file|post|postfile|postclose|putexcel|putdocx|graph|gr\s+export|twoway|scatter|line|bar|histogram|kdensity|plot|tsset|xtset|tsline|tsplot|corr|pwcorr|tabulate|tab|tab1|tab2|tabstat|table|ttest|mean|proportion|pctile|summarize|sum|count|distinct|regress|areg|xtreg|ivregress|ivreg2|reghdfe|logit|probit|mlogit|ologit|oprobit|poisson|nbreg|qreg|sqreg|rq|anova|manova|mixed|meqrlogit|gsem|sem|margins|marginsplot|predict|estat|test|lincom|nlcom|svy|svyset|mi|mi\s+set|mi\s+register|mi\s+impute|mi\s+estimate|estimates|eststo|esttab|estout|parmest|xtdescribe|xtsum|xttest|xtoverid|xtserial)\b", RegexOptions.IgnoreCase, "keyword"),

                        // Line continuation operator ///
                        (@"\/\/\/\s*$", RegexOptions.Multiline, "selector")
                    }
                },
            };
        }

        private void InitializeAliases()
        {
            languageAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "js", "javascript" },
                { "c#", "csharp" },
                { "cs", "csharp" },
                { "py", "python" },
                { "m", "matlab" },
                { "css", "css" },
                { "htm", "html" },
                { "R", "r" },
                { "Fortran", "fortran" },
                { "f90", "fortran" },
                { "f95", "fortran" },
                { "f03", "fortran" },
                { "f08", "fortran" },
                { "f77", "fortran" },
                { "for", "fortran" },
                { "f", "fortran" },
                // Stata aliases
                { "stata", "stata" },
                { "do", "stata" },
                { "do-file", "stata" },
                { "ado", "stata" }
            };
        }

        private int GetActualPosition(string text, int position)
        {
            // 计算到指定位置前的换行符数量
            int newlineCount = text.Substring(0, position).Count(c => c == '\n' || c == '\r');
            // PowerPoint中每个换行符只算一个字符，而在字符串中\r\n算两个字符
            int adjustment = text.Substring(0, position).Count(c => c == '\r');
            return position - adjustment;
        }

        public void ApplyHighlighting(PowerPoint.Shape textBox, string code, string language)
        {
            // Normalize language name
            string normalizedLanguage = language.ToLower();
            if (languageAliases.ContainsKey(normalizedLanguage))
            {
                normalizedLanguage = languageAliases[normalizedLanguage];
            }

            if (!languagePatterns.ContainsKey(normalizedLanguage))
                return;

            processedRanges.Clear();
            var patterns = languagePatterns[normalizedLanguage];

            // 收集所有匹配项
            var allMatches = new List<(Match match, string type, int priority)>();

            foreach (var (pattern, options, type) in patterns)
            {
                var regex = new Regex(pattern, options);
                var matches = regex.Matches(code);

                // 设置优先级：注释最高优先级，然后是字符串
                int priority = type == "comment" ? 0 : (type == "string" ? 1 : 2);

                foreach (Match match in matches)
                {
                    allMatches.Add((match, type, priority));
                }
            }

            // 过滤掉位于字符串内部的注释（!）

            // 先获取所有字符串区间
            var stringIntervals = allMatches
                .Where(m => m.type == "string")
                .Select(m => (start: m.match.Index, end: m.match.Index + m.match.Length))
                .ToList();

            if (stringIntervals.Count > 0)
            {
                allMatches = allMatches
                    .Where(m => !(m.type == "comment" && stringIntervals.Any(si => m.match.Index >= si.start && m.match.Index < si.end)))
                    .ToList();
            }

            // 按优先级排序，优先级数字越小越优先
            allMatches.Sort((a, b) => a.priority.CompareTo(b.priority));

            // 记录已处理的字符范围
            var processedIntervals = new List<(int start, int end)>();

            foreach (var (match, type, priority) in allMatches)
            {
                try
                {
                    int matchStart = match.Index;
                    int matchEnd = match.Index + match.Length;

                    // 检查是否与已处理的区域重叠
                    bool hasOverlap = false;
                    foreach (var (start, end) in processedIntervals)
                    {
                        if (matchStart < end && matchEnd > start)
                        {
                            hasOverlap = true;
                            break;
                        }
                    }

                    if (hasOverlap)
                        continue;

                    // 计算实际的开始位置和长度
                    int actualStart = GetActualPosition(code, matchStart);
                    int actualLength = GetActualPosition(code, matchEnd) - actualStart;

                    var range = textBox.TextFrame.TextRange.Characters(
                        actualStart + 1,
                        actualLength
                    );
                    range.Font.Color.RGB = ColorTranslator.ToOle(themeColors[type]);

                    // 记录已处理的区间
                    processedIntervals.Add((matchStart, matchEnd));
                }
                catch (Exception)
                {
                    continue;
                }
            }
        }
    }
}
