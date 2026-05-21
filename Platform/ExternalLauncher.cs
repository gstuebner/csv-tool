using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace CsvTool.Platform
{
    public static class ExternalLauncher
    {
        public static bool LaunchExcel(string filePath)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                // 1. Try via PATH "excel"
                if (StartProcess("excel", $"\"{filePath}\"")) return true;

                // 2. Search Common Paths
                string[] paths = {
                    @"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
                    @"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
                };

                foreach (var p in paths)
                {
                    if (File.Exists(p))
                    {
                        if (StartProcess(p, $"\"{filePath}\"")) return true;
                    }
                }
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                return StartProcess("xdg-open", $"\"{filePath}\"");
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                return StartProcess("open", $"\"{filePath}\"");
            }

            return false;
        }

        public static bool LaunchLibreOffice(string filePath)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                // 1. Try via PATH "scalc" or "soffice"
                if (StartProcess("scalc", $"\"{filePath}\"")) return true;

                // 2. Common Paths
                string[] paths = {
                    @"C:\Program Files\LibreOffice\program\scalc.exe",
                    @"C:\Program Files (x86)\LibreOffice\program\scalc.exe"
                };

                foreach (var p in paths)
                {
                    if (File.Exists(p))
                    {
                        if (StartProcess(p, $"\"{filePath}\"")) return true;
                    }
                }
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                // Try scalc directly, then libreoffice --calc, then soffice --calc, then xdg-open
                if (StartProcess("scalc", $"\"{filePath}\"")) return true;
                if (StartProcess("libreoffice", $"--calc \"{filePath}\"")) return true;
                if (StartProcess("soffice", $"--calc \"{filePath}\"")) return true;
                return StartProcess("xdg-open", $"\"{filePath}\"");
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                return StartProcess("open", $"-a \"LibreOffice\" \"{filePath}\"") || StartProcess("open", $"\"{filePath}\"");
            }

            return false;
        }

        private static bool StartProcess(string exe, string args)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = exe,
                    Arguments = args,
                    UseShellExecute = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                });
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
