using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace SlideSCI
{
    public class LatexToSvgConverter
    {
        private readonly string _nodeExecutable;
        private readonly string _scriptPath;
        private readonly string _workingDirectory;

        public LatexToSvgConverter(string nodeExecutable = "node")
        {
            _nodeExecutable = string.IsNullOrWhiteSpace(nodeExecutable) ? "node" : nodeExecutable;

            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory ?? AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            if (string.IsNullOrEmpty(baseDirectory))
            {
                baseDirectory = AppDomain.CurrentDomain.SetupInformation?.ApplicationBase ?? string.Empty;
            }

            _workingDirectory = Path.Combine(baseDirectory, "latex-converter");
            _scriptPath = Path.Combine(_workingDirectory, "latex-to-svg.js");
        }

        public string ConvertLatexToSvg(string latexCode)
        {
            if (string.IsNullOrWhiteSpace(latexCode))
            {
                throw new ArgumentException("LaTeX 公式不能为空。", nameof(latexCode));
            }

            if (!File.Exists(_scriptPath))
            {
                throw new FileNotFoundException(
                    "未找到 LaTeX 转换脚本，请确认插件目录下的 latex-converter 文件夹已部署。",
                    _scriptPath
                );
            }

            var processStartInfo = new ProcessStartInfo
            {
                FileName = _nodeExecutable,
                Arguments = $"\"{_scriptPath}\"",
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
                WorkingDirectory = _workingDirectory,
            };

            try
            {
                using (var process = Process.Start(processStartInfo))
                {
                    if (process == null)
                    {
                        throw new InvalidOperationException("无法启动 Node.js 进程，请确认已安装 Node.js 并可在系统 PATH 中访问。");
                    }

                    using (var writer = new StreamWriter(process.StandardInput.BaseStream, Encoding.UTF8))
                    {
                        writer.Write(latexCode);
                        writer.Flush();
                    }

                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();

                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        string message = string.IsNullOrWhiteSpace(error)
                            ? "LaTeX 转换失败，Node.js 返回非零状态码。"
                            : error;
                        throw new InvalidOperationException(message);
                    }

                    if (string.IsNullOrWhiteSpace(output))
                    {
                        throw new InvalidOperationException("LaTeX 转换结果为空，请检查输入内容是否有效。");
                    }

                    return output.Trim();
                }
            }
            catch (FileNotFoundException)
            {
                throw;
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"LaTeX 转 SVG 处理时出现异常：{ex.Message}", ex);
            }
        }
    }
}
