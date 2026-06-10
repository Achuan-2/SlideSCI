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
                        throw new InvalidOperationException($"无法启动 Node.js 进程。请检查以下设置：\n" +
                            $"1. 确认已安装 Node.js 并可在系统 PATH 中访问\n" +
                            $"2. 下载地址：https://nodejs.org/\n" +
                            $"3. 安装完成后重启 PowerPoint");
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
                        string message;
                        if (string.IsNullOrWhiteSpace(error))
                        {
                            message = "LaTeX 转换失败，Node.js 返回非零状态码。";
                        }
                        else
                        {
                            // 检查常见的错误类型并提供具体指导
                            if (error.Contains("缺少必要的依赖包") || error.Contains("mathjax-full"))
                            {
                                message = $"缺少 Node.js 依赖包。请按以下步骤安装：\n" +
                                    $"1. 打开命令提示符\n" +
                                    $"2. 导航到插件目录：{_workingDirectory}\n" +
                                    $"3. 然后运行：npm install\n" +
                                    $"4. 安装完成后重试\n\n";
                            }
                            else
                            {
                                message = error;
                            }
                        }
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
            catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 2)
            {
                throw new InvalidOperationException($"找不到 Node.js 可执行文件。请检查以下设置：\n" +
                    $"1. 确认已安装 Node.js（下载地址：https://nodejs.org/）\n" +
                    $"2. 确认 Node.js 已添加到系统环境变量 PATH 中\n" +
                    $"3. 安装完成后请重启 PowerPoint\n" +
                    $"4. 如果问题仍然存在，请在命令提示符中运行 'node --version' 确认安装是否成功");
            }
            catch (Exception ex)
            {
                string errorMessage = ex.Message;
                
                // 检查是否为依赖缺失错误
                if (errorMessage.Contains("Cannot find module") || errorMessage.Contains("mathjax-full"))
                {
                    throw new InvalidOperationException($"缺少必要的 Node.js 依赖包。请按以下步骤安装：\n" +
                        $"1. 打开命令提示符（以管理员身份运行）\n" +
                        $"2. 导航到插件目录：{_workingDirectory}\n" +
                        $"3. 然后运行：npm install\n" +
                        $"4. 安装完成后重试LaTeX转换功能\n\n" +
                        $"原始错误信息：{errorMessage}");
                }
                
                throw new InvalidOperationException($"LaTeX 转 SVG 处理时出现异常：{errorMessage}\n\n" +
                    $"如果是首次使用，请确认：\n" +
                    $"1. 已安装 Node.js（https://nodejs.org/）\n" +
                    $"2. 已在插件目录 {_workingDirectory} 中运行 'npm install' 安装依赖", ex);
            }
        }
    }
}
