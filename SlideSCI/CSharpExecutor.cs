using System;
using System.CodeDom.Compiler;
using System.IO;
using System.Reflection;
using System.Text;
using Microsoft.CSharp;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    public static class CSharpExecutor
    {
        public static string Execute(string code, PowerPoint.Application app)
        {
            using (var provider = new CSharpCodeProvider())
            {
                var parameters = new CompilerParameters();
                parameters.GenerateInMemory = true;
                
                // Add standard references
                parameters.ReferencedAssemblies.Add("System.dll");
                parameters.ReferencedAssemblies.Add("System.Core.dll");
                parameters.ReferencedAssemblies.Add("System.Drawing.dll");
                parameters.ReferencedAssemblies.Add("System.Windows.Forms.dll");
                
                // Add PowerPoint Interop and Office Core references dynamically using GAC or local fallback
                try
                {
                    bool pptLoaded = false;
                    try
                    {
                        var pptAssembly = Assembly.Load("Microsoft.Office.Interop.PowerPoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c");
                        if (!string.IsNullOrEmpty(pptAssembly.Location))
                        {
                            parameters.ReferencedAssemblies.Add(pptAssembly.Location);
                            pptLoaded = true;
                        }
                    }
                    catch { }

                    if (!pptLoaded)
                    {
                        string windir = Environment.GetEnvironmentVariable("windir") ?? @"C:\Windows";
                        string gacPath = Path.Combine(windir, @"assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.PowerPoint.dll");
                        if (File.Exists(gacPath))
                        {
                            parameters.ReferencedAssemblies.Add(gacPath);
                        }
                        else
                        {
                            parameters.ReferencedAssemblies.Add(typeof(PowerPoint.Application).Assembly.Location);
                        }
                    }

                    bool officeLoaded = false;
                    try
                    {
                        var officeAssembly = Assembly.Load("Office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c");
                        if (!string.IsNullOrEmpty(officeAssembly.Location))
                        {
                            parameters.ReferencedAssemblies.Add(officeAssembly.Location);
                            officeLoaded = true;
                        }
                    }
                    catch { }

                    if (!officeLoaded)
                    {
                        string windir = Environment.GetEnvironmentVariable("windir") ?? @"C:\Windows";
                        string gacPath = Path.Combine(windir, @"assembly\GAC_MSIL\Office\15.0.0.0__71e9bce111e9429c\Office.dll");
                        if (File.Exists(gacPath))
                        {
                            parameters.ReferencedAssemblies.Add(gacPath);
                        }
                        else
                        {
                            parameters.ReferencedAssemblies.Add(typeof(Office.MsoTriState).Assembly.Location);
                        }
                    }
                }
                catch (Exception ex)
                {
                    return $"Failed to load PowerPoint assembly references: {ex.Message}";
                }

                string source = $@"
using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

public class ScriptRunner
{{
    public static void Run(PowerPoint.Application Application)
    {{
        {code}
    }}
}}";

                var results = provider.CompileAssemblyFromSource(parameters, source);
                if (results.Errors.HasErrors)
                {
                    var errors = new StringBuilder();
                    errors.AppendLine("Compilation Errors:");
                    foreach (CompilerError err in results.Errors)
                    {
                        errors.AppendLine($"- Line {err.Line - 13}: Error ({err.ErrorNumber}): {err.ErrorText}");
                    }
                    return errors.ToString();
                }

                try
                {
                    var assembly = results.CompiledAssembly;
                    var type = assembly.GetType("ScriptRunner");
                    var method = type.GetMethod("Run", BindingFlags.Static | BindingFlags.Public);
                    
                    // Run the script
                    method.Invoke(null, new object[] { app });
                    return "Success";
                }
                catch (TargetInvocationException ex)
                {
                    var inner = ex.InnerException ?? ex;
                    return $"Runtime Error: {inner.Message}\nStack Trace:\n{inner.StackTrace}";
                }
                catch (Exception ex)
                {
                    return $"Execution Error: {ex.Message}\nStack Trace:\n{ex.StackTrace}";
                }
            }
        }
    }
}
