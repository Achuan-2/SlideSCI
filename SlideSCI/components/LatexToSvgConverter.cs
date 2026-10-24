using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Jint;

namespace SlideSCI
{
    /// <summary>Runs the bundled MathJax SVG renderer in-process, without Node.js.</summary>
    public sealed class LatexToSvgConverter : IDisposable
    {
        private const string ScriptResourceName = "SlideSCI.MathJax.svg.min.js";
        private readonly object _sync = new object();
        private readonly CancellationTokenSource _shutdown = new CancellationTokenSource();
        private Engine _engine;
        private Task _warmupTask;
        private int _disposeStarted;

        public Task WarmUpAsync()
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                return _warmupTask ?? (_warmupTask = Task.Run(() =>
                {
                    try { ConvertLatexToSvg("x"); }
                    catch (Exception ex) { Debug.WriteLine("SlideSCI MathJax warmup: " + ex.Message); }
                }));
            }
        }

        public string ConvertLatexToSvg(string latexCode)
        {
            if (string.IsNullOrWhiteSpace(latexCode))
                throw new ArgumentException("LaTeX 公式不能为空。", nameof(latexCode));

            string text = latexCode.Trim();
            bool display = text.StartsWith(@"\[") || text.StartsWith("$$")
                || text.Contains(@"\\")
                || Regex.IsMatch(text, @"\\begin\s*\{(align|equation|gather|multline|cases|matrix)");
            string source = NormalizeLatex(text);
            if (string.IsNullOrWhiteSpace(source))
                throw new ArgumentException("LaTeX 公式不能为空。", nameof(latexCode));

            lock (_sync)
            {
                ThrowIfDisposed();
                EnsureInitialized();
                string result = _engine.Invoke("slideSciToSvg", source, display).ToString();
                var xml = new XmlDocument { XmlResolver = null };
                xml.LoadXml(result);
                if (xml.DocumentElement?.LocalName != "svg"
                    || xml.DocumentElement.NamespaceURI != "http://www.w3.org/2000/svg")
                    throw new InvalidOperationException("MathJax 未生成有效 SVG。");
                return result;
            }
        }

        private static string NormalizeLatex(string text)
        {
            if (text.Length >= 4 && ((text.StartsWith("$$") && text.EndsWith("$$"))
                || (text.StartsWith(@"\(") && text.EndsWith(@"\)"))
                || (text.StartsWith(@"\[") && text.EndsWith(@"\]"))))
                return text.Substring(2, text.Length - 4).Trim();
            return text.Length >= 2 && text.StartsWith("$") && text.EndsWith("$")
                ? text.Substring(1, text.Length - 2).Trim() : text;
        }

        private void EnsureInitialized()
        {
            if (_engine != null) return;
            var engine = new Engine(options => options
                .TimeoutInterval(TimeSpan.FromSeconds(15))
                .CancellationToken(_shutdown.Token)
                .LimitMemory(256 * 1024 * 1024)
                .LimitRecursion(512));
            try
            {
                using (Stream stream = typeof(LatexToSvgConverter).Assembly.GetManifestResourceStream(ScriptResourceName))
                {
                    if (stream == null) throw new InvalidOperationException("插件缺少内置 MathJax 资源，请重新安装完整插件。");
                    using (var reader = new StreamReader(stream)) { engine.Execute(reader.ReadToEnd(), "mathjax-svg.min.js"); }
                }
                _engine = engine;
            }
            catch { engine.Dispose(); throw; }
        }

        private void ThrowIfDisposed()
        {
            if (_disposeStarted != 0) throw new ObjectDisposedException(nameof(LatexToSvgConverter));
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
