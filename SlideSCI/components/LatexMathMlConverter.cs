using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Jint;
using Newtonsoft.Json;

namespace SlideSCI
{
    /// <summary>Converts LaTeX to MathML inside the add-in, without Node.js or a browser.</summary>
    public sealed class LatexMathMlConverter : IDisposable
    {
        private const string ScriptResourceName = "SlideSCI.KaTeX.katex.min.js";
        private const string MathMlNamespace = "http://www.w3.org/1998/Math/MathML";
        private const string BridgeScript = @"
            function slideSciToMathMl(json) {
                var macros = Object.create(null);
                return JSON.stringify(JSON.parse(json).map(function (latex) {
                    return katex.renderToString(latex, {
                        output: 'mathml', displayMode: false, throwOnError: true,
                        strict: 'ignore', trust: false, maxExpand: 1000,
                        macros: macros, globalGroup: true
                    });
                }));
            }";

        private readonly object _sync = new object();
        private readonly CancellationTokenSource _shutdown = new CancellationTokenSource();
        private readonly string _script;
        private Engine _engine;
        private Task _warmupTask;
        private int _disposeStarted;

        // Supplying script text allows isolated tests; production uses the embedded resource.
        public LatexMathMlConverter(string script = null)
        {
            _script = script;
        }

        public Task WarmUpAsync()
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                return _warmupTask ?? (_warmupTask = Task.Run(() =>
                {
                    try { ConvertLatexToMathMl("x"); }
                    catch (Exception ex) { Debug.WriteLine("SlideSCI KaTeX warmup: " + ex.Message); }
                }));
            }
        }

        public string ConvertLatexToMathMl(string latexCode)
        {
            return ConvertLatexToMathMl(new[] { latexCode })[0];
        }

        public string[] ConvertLatexToMathMl(string[] latexCodes)
        {
            if (latexCodes == null || latexCodes.Length == 0) return new string[0];
            if (Array.Exists(latexCodes, string.IsNullOrWhiteSpace))
            {
                throw new ArgumentException("LaTeX 公式不能为空。", nameof(latexCodes));
            }

            lock (_sync)
            {
                ThrowIfDisposed();
                EnsureInitialized();
                var normalized = new string[latexCodes.Length];
                for (int i = 0; i < latexCodes.Length; i++) normalized[i] = NormalizeLatex(latexCodes[i]);
                string output = _engine.Invoke("slideSciToMathMl", JsonConvert.SerializeObject(normalized)).ToString();
                string[] results = JsonConvert.DeserializeObject<string[]>(output);
                if (results == null || results.Length != latexCodes.Length)
                {
                    throw new InvalidOperationException("KaTeX 返回的公式数量不正确。");
                }
                for (int i = 0; i < results.Length; i++)
                {
                    results[i] = ExtractMathMl(results[i]);
                }
                return results;
            }
        }

        private static string NormalizeLatex(string latex)
        {
            string text = latex.Trim();
            if (text.Length >= 4 && ((text.StartsWith("$$") && text.EndsWith("$$"))
                || (text.StartsWith(@"\(") && text.EndsWith(@"\)"))
                || (text.StartsWith(@"\[") && text.EndsWith(@"\]"))))
            {
                return text.Substring(2, text.Length - 4).Trim();
            }
            return text.Length >= 2 && text.StartsWith("$") && text.EndsWith("$")
                ? text.Substring(1, text.Length - 2).Trim() : text;
        }

        private void EnsureInitialized()
        {
            if (_engine != null) return;
            Engine engine = new Engine(options => options
                .TimeoutInterval(TimeSpan.FromSeconds(5))
                .CancellationToken(_shutdown.Token)
                .LimitMemory(64 * 1024 * 1024)
                .LimitRecursion(256));
            try
            {
                engine.Execute(_script ?? ReadBundledScript(), "katex.min.js");
                engine.Execute(BridgeScript, "slidesci-katex.js");
                _engine = engine;
            }
            catch
            {
                engine.Dispose();
                throw;
            }
        }

        private static string ReadBundledScript()
        {
            using (Stream stream = typeof(LatexMathMlConverter).Assembly.GetManifestResourceStream(ScriptResourceName))
            {
                if (stream == null) throw new InvalidOperationException("插件缺少内置 KaTeX 资源，请重新安装完整插件。");
                using (var reader = new StreamReader(stream)) { return reader.ReadToEnd(); }
            }
        }

        private static string ExtractMathMl(string markup)
        {
            var document = new XmlDocument { XmlResolver = null };
            document.LoadXml(markup);
            XmlNode math = document.GetElementsByTagName("math", MathMlNamespace).Item(0);
            if (math == null) throw new InvalidOperationException("KaTeX 未生成有效 MathML。");
            var namespaces = new XmlNamespaceManager(document.NameTable);
            namespaces.AddNamespace("m", MathMlNamespace);
            // Office needs the presentation tree, not KaTeX's wrapper or source annotation.
            foreach (XmlNode annotation in math.SelectNodes(".//m:annotation | .//m:annotation-xml", namespaces))
            {
                annotation.ParentNode.RemoveChild(annotation);
            }
            return math.OuterXml;
        }

        private void ThrowIfDisposed()
        {
            if (_disposeStarted != 0) throw new ObjectDisposedException(nameof(LatexMathMlConverter));
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposeStarted, 1) != 0) return;
            _shutdown.Cancel();
            lock (_sync) { _engine?.Dispose(); _engine = null; }
            _shutdown.Dispose();
        }
    }
}
