using System;
using System.IO;

namespace CertificateGenerator
{
    public static class Logger
    {
        private static string _logPath;

        public static void Init(string baseDir)
        {
            _logPath = Path.Combine(baseDir, "application.log");
            try { File.WriteAllText(_logPath, $"--- Process Started: {DateTime.Now} ---\n"); } catch { }
        }

        public static void Log(string message)
        {
            string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {message}";
            Console.WriteLine(entry);
            try
            {
                if (!string.IsNullOrEmpty(_logPath))
                    File.AppendAllText(_logPath, entry + Environment.NewLine);
            }
            catch { }
        }
    }
}