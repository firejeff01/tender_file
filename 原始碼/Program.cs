using System;
using System.Windows.Forms;

namespace TenderDocGen
{
    /// <summary>程式進入點。一般模式開啟 GUI；「/selftest」為開發自我測試模式。</summary>
    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            if (args.Length > 0 && string.Equals(args[0], "/selftest", StringComparison.OrdinalIgnoreCase))
            {
                return SelfTest.Run(args);
            }

            // 任何未被攔截的例外都以可讀對話框呈現，避免非技術使用者看到 .NET 崩潰視窗
            Application.ThreadException += delegate(object s, System.Threading.ThreadExceptionEventArgs e)
            {
                ShowUnhandled(e.Exception);
            };
            AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e)
            {
                ShowUnhandled(e.ExceptionObject as Exception);
            };

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            return 0;
        }

        static void ShowUnhandled(Exception ex)
        {
            string msg = ex != null ? ex.Message : "未知的錯誤";
            MessageBox.Show(
                "程式發生非預期的錯誤：\n\n" + msg +
                "\n\n請將此訊息告知資訊人員。程式將繼續執行，但剛才的操作可能未完成。",
                "發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
