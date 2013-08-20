using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace SOSync
{
    class Logger
    {
        public const string LEVEL_DEBUG = "DEBUG";
        public const string LEVEL_INFO = "INFO";
        public const string LEVEL_WARNING = "WARNING";
        public const string LEVEL_ERROR = "ERROR";

        private static ListBox _lb = null;
        private static volatile int _indent = 0;

        public static void addListBox(ListBox lb)
        {
            _lb = lb;
        }

        public static string Log(string level, string logMessage)
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");

            string fileName = "log-" + today + ".txt";
            StreamWriter w = null;
            if (File.Exists(fileName))
            {
                w = File.AppendText(fileName);
            }
            else
            {
                w = File.CreateText(fileName);
            }

            string message = String.Format("[{0}][{1}][{2}] {3}", DateTime.Now.ToString(),
                level, _indent, logMessage);

            w.WriteLine(message);
            w.Flush();
            w.Close();

            if (level != LEVEL_DEBUG)
            {
                LogToListBox(message);
            }
            return message;
        }

        public static void LogToListBox(string message) {
            if (_lb != null) {
                if (_lb.InvokeRequired) {
                    _lb.BeginInvoke((Action)(() => LogToListBox(message)));
                } else {
                    _lb.Items.Insert(0, message);
                }
            }
        }
        public static void Indent()
        {
            _indent++;
        }
        public static void Dedent()
        {
            _indent--;
        }
        
        public static IDisposable Scope(string message, string level = LEVEL_INFO)
        {
            return new LoggerIndent(message, level);
        }

        class LoggerIndent : IDisposable
        {
            private string message;
            private string level;
            public LoggerIndent(string message, string level = LEVEL_INFO) {
                this.message = message;
                this.level = level;
                Logger.Indent();
                Logger.Log(level, "[START] " + message);
            }
            public void Dispose()
            {
                Logger.Log(level, "[DONE] " + message);
                Logger.Dedent();
            }
        }
    }
}
