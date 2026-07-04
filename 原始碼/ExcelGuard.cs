using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace TenderDocGen
{
    /// <summary>
    /// 更新 Excel 前確保「標案資料.xlsx」沒被 Excel 佔用。
    /// 溫和關閉（Process.CloseMainWindow）：Excel 若有未存變動會跳它自己的存檔詢問，
    /// 由使用者處理，程式不強殺、不丟資料；再輪詢檔案是否釋出。
    /// </summary>
    static class ExcelGuard
    {
        /// <summary>回傳是否可繼續（true = 檔案可寫，可安全更新）。</summary>
        public static bool EnsureClosed(string xlsxPath, IWin32Window owner)
        {
            if (!File.Exists(xlsxPath)) return true;   // 沒有檔案，交給後續流程處理
            if (IsWritable(xlsxPath)) return true;     // 沒被佔用

            DialogResult dr = MessageBox.Show(owner,
                "接下來需要更新「標案資料.xlsx」，必須先關閉 Excel。\n\n" +
                "請先確認 Excel 內的資料已存檔。按「確定」後程式會關閉 Excel 並繼續；\n" +
                "若 Excel 詢問是否儲存，請依需要選擇。",
                "需要關閉 Excel", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (dr != DialogResult.OK) return false;

            // 溫和要求所有 Excel 視窗關閉
            foreach (Process p in SafeGetExcelProcesses())
            {
                try { if (!p.HasExited && p.MainWindowHandle != IntPtr.Zero) p.CloseMainWindow(); }
                catch { }
            }

            // 輪詢檔案釋出（最多約 15 秒，容納使用者在存檔詢問上操作）
            for (int i = 0; i < 60; i++)
            {
                if (IsWritable(xlsxPath)) return true;
                Thread.Sleep(250);
            }

            if (IsWritable(xlsxPath)) return true;
            MessageBox.Show(owner,
                "Excel 尚未關閉（或檔案仍被佔用），操作已取消。\n請手動關閉 Excel 後再試一次。",
                "操作取消", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return false;
        }

        /// <summary>能否以獨占方式開啟（＝沒被別的程式鎖住）。</summary>
        static bool IsWritable(string path)
        {
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    return true;
            }
            catch (IOException) { return false; }
            catch (UnauthorizedAccessException) { return false; }
        }

        static Process[] SafeGetExcelProcesses()
        {
            try { return Process.GetProcessesByName("EXCEL"); }
            catch { return new Process[0]; }
        }
    }
}
