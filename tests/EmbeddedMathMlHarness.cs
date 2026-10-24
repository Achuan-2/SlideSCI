using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Xml;
using Newtonsoft.Json;
using SlideSCI;

internal static class EmbeddedMathMlHarness
{
    private static string repoPath;

    public static int Main(string[] args)
    {
        try
        {
            repoPath = Path.GetFullPath(args[0]);
            AppDomain.CurrentDomain.AssemblyResolve += ResolveDependency;
            // The harness exercises the actual .NET Framework target and cannot find Node on PATH.
            Environment.SetEnvironmentVariable("PATH", string.Empty);
            Run();
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }

    private static Assembly ResolveDependency(object sender, ResolveEventArgs args)
    {
        var paths = new Dictionary<string, string>
        {
            ["Jint"] = "Jint.4.16.0/lib/net462/Jint.dll",
            ["Acornima"] = "Acornima.1.7.0/lib/net462/Acornima.dll",
            ["Newtonsoft.Json"] = "Newtonsoft.Json.13.0.3/lib/net45/Newtonsoft.Json.dll",
            ["System.Memory"] = "System.Memory.4.5.5/lib/net461/System.Memory.dll",
            ["System.Buffers"] = "System.Buffers.4.5.1/lib/net461/System.Buffers.dll",
            ["System.Numerics.Vectors"] = "System.Numerics.Vectors.4.5.0/lib/net46/System.Numerics.Vectors.dll",
            ["System.Runtime.CompilerServices.Unsafe"] = "System.Runtime.CompilerServices.Unsafe.6.1.0/lib/net462/System.Runtime.CompilerServices.Unsafe.dll"
        };
        string relative;
        return paths.TryGetValue(new AssemblyName(args.Name).Name, out relative)
            ? Assembly.LoadFrom(Path.Combine(repoPath, "packages", relative)) : null;
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void Run()
    {
        Assert(typeof(LatexMathMlConverter).Assembly.GetManifestResourceNames().Contains("SlideSCI.KaTeX.katex.min.js"), "KaTeX was not embedded.");
        var timings = new List<object>();
        var timer = Stopwatch.StartNew();
        var converter = new LatexMathMlConverter();
        try
        {
            Task warmup = converter.WarmUpAsync();
            timings.Add(new { Stage = "schedule_warmup", Milliseconds = timer.Elapsed.TotalMilliseconds });
            warmup.GetAwaiter().GetResult();
            timings.Add(new { Stage = "ready", Milliseconds = timer.Elapsed.TotalMilliseconds });
            string[] formulas = { "L^{(v)}", "E = mc^2", "S_a^{(v)}", @"\frac{a_1}{b^2}", @"\begin{pmatrix}a&b\\c&d\end{pmatrix}" };
            string[] structures = { "msup", "msup", "msubsup", "mfrac", "mtable" };
            for (int i = 0; i < formulas.Length; i++)
            {
                timer.Restart();
                string math = converter.ConvertLatexToMathMl(formulas[i]);
                timings.Add(new { Stage = formulas[i], Milliseconds = timer.Elapsed.TotalMilliseconds });
                Assert(Count(math, structures[i]) > 0, "Missing structure: " + formulas[i]);
                Assert(Count(math, "annotation") == 0 && !math.Contains("<span"), "KaTeX wrapper leaked into MathML.");
            }
            Assert(Count(converter.ConvertLatexToMathMl(@"\(L^{(v)}\)"), "msup") == 1, "Delimited input failed.");
            var chinese = new XmlDocument();
            chinese.LoadXml(converter.ConvertLatexToMathMl(@"\text{中文}"));
            Assert(chinese.InnerText.Contains("中文"), "Chinese text was lost.");
            AssertFails(() => converter.ConvertLatexToMathMl(@"\frac{a}{"), "Invalid TeX was accepted.");
            Assert(Count(converter.ConvertLatexToMathMl("x^2"), "msup") == 1, "Engine did not recover after invalid TeX.");
            string[] defined = converter.ConvertLatexToMathMl(new[] { @"\newcommand{\onlyhere}{x}\onlyhere", @"\onlyhere" });
            Assert(defined[1].Contains(">x</mi>"), "Macros failed inside one batch.");
            AssertFails(() => converter.ConvertLatexToMathMl(@"\onlyhere"), "Macros leaked into a later request.");
            AssertFails(() => converter.ConvertLatexToMathMl(@"\def\loop{\loop}\loop"), "Unbounded macro expansion was accepted.");
            Task<string>[] concurrent = Enumerable.Range(1, 4)
                .Select(index => Task.Run(() => converter.ConvertLatexToMathMl("x_{" + index + "}"))).ToArray();
            string[] outputs = Task.WhenAll(concurrent).GetAwaiter().GetResult();
            for (int i = 0; i < outputs.Length; i++) Assert(outputs[i].Contains(">" + (i + 1) + "</mn>"), "Concurrent responses were mixed.");
            AssertFails(() => converter.ConvertLatexToMathMl(@"\begin{prooftree}\AxiomC{$A$}\end{prooftree}"), "Unsupported environments must report an error.");
            Assert(Count(converter.ConvertLatexToMathMl("y^3"), "msup") == 1, "Unsupported input poisoned the engine.");
            File.WriteAllText(Path.Combine(repoPath, "artifacts", "diagnostics", "embedded-katex-timings.json"), JsonConvert.SerializeObject(timings, Newtonsoft.Json.Formatting.Indented));
            Console.WriteLine(JsonConvert.SerializeObject(timings, Newtonsoft.Json.Formatting.Indented));
        }
        finally { converter.Dispose(); converter.Dispose(); }
        AssertFails(() => converter.ConvertLatexToMathMl("x"), "Disposed engine was reused.");
        var cancelling = new LatexMathMlConverter();
        Task pending = cancelling.WarmUpAsync();
        cancelling.Dispose();
        pending.GetAwaiter().GetResult();
        Console.WriteLine("Passed .NET Framework conversion with empty PATH, embedded resources, formulas, UTF-8, errors, macros, concurrency, expansion limits, and disposal.");
        VerifySvg();
    }

    private static void VerifySvg()
    {
        var converter = new LatexToSvgConverter();
        var timings = new List<object>();
        string directory = Path.Combine(repoPath, "artifacts", "diagnostics", "embedded-svg");
        Directory.CreateDirectory(directory);
        string[] formulas = {
            @"\(L^{(v)}\)", @"\frac{a_1}{b^2}", @"\begin{pmatrix}a&b\\c&d\end{pmatrix}",
            @"\begin{prooftree}\AxiomC{$A$}\UnaryInfC{$B$}\end{prooftree}",
            @"\ce{H2O}", @"\begin{align}a&=b+c\\d&=e\end{align}", @"\text{中文}"
        };
        try
        {
            for (int i = 0; i < formulas.Length; i++)
            {
                var timer = Stopwatch.StartNew();
                string svg = converter.ConvertLatexToSvg(formulas[i]);
                timings.Add(new { Formula = formulas[i], Milliseconds = timer.Elapsed.TotalMilliseconds });
                var xml = new XmlDocument { XmlResolver = null };
                xml.LoadXml(svg);
                Assert(xml.DocumentElement.LocalName == "svg", "Missing SVG root.");
                Assert(xml.GetElementsByTagName("path").Count > 0 || i == 6, "Missing embedded glyph paths.");
                Assert(!svg.Contains("currentColor") && !svg.Contains("data-mjx-error="), "Invalid SVG output.");
                if (i == 3) Assert(xml.GetElementsByTagName("line").Cast<XmlElement>()
                    .All(line => line.GetAttribute("stroke-width") == "70"), "Proof-tree lines rely on missing external CSS.");
                File.WriteAllText(Path.Combine(directory, "formula-" + i + ".svg"), svg);
            }
            AssertFails(() => converter.ConvertLatexToSvg(@"\frac{a}{"), "Invalid SVG source was accepted.");
            AssertFails(() => converter.ConvertLatexToSvg(@"\unknownSlideSciCommand"), "Unknown command was accepted.");
            AssertFails(() => converter.ConvertLatexToSvg("$$"), "Empty delimited source was accepted.");
            converter.ConvertLatexToSvg(@"\newcommand{\onlyhere}{x}\onlyhere");
            AssertFails(() => converter.ConvertLatexToSvg(@"\onlyhere"), "SVG macros leaked into the next call.");
            Task<string>[] concurrent = Enumerable.Range(1, 3)
                .Select(index => Task.Run(() => converter.ConvertLatexToSvg("x^{" + index + "}"))).ToArray();
            string[] outputs = Task.WhenAll(concurrent).GetAwaiter().GetResult();
            Assert(outputs.Distinct().Count() == 3, "Concurrent SVG responses were mixed.");
            converter.ConvertLatexToSvg("y^3");
        }
        finally { converter.Dispose(); converter.Dispose(); }
        AssertFails(() => converter.ConvertLatexToSvg("x"), "Disposed SVG engine was reused.");
        var cancelling = new LatexToSvgConverter();
        Task pending = cancelling.WarmUpAsync();
        cancelling.Dispose();
        pending.GetAwaiter().GetResult();
        string json = JsonConvert.SerializeObject(timings, Newtonsoft.Json.Formatting.Indented);
        File.WriteAllText(Path.Combine(directory, "timings.json"), json);
        Console.WriteLine(json);
        Console.WriteLine("Passed embedded MathJax SVG on .NET Framework with empty PATH, advanced formulas, error recovery, macro isolation, concurrency, and disposal.");
    }

    private static int Count(string math, string name)
    {
        var xml = new XmlDocument();
        xml.LoadXml(math);
        return xml.GetElementsByTagName(name, "http://www.w3.org/1998/Math/MathML").Count;
    }

    private static void AssertFails(Action action, string message)
    {
        try { action(); } catch { return; }
        throw new InvalidOperationException(message);
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
}
