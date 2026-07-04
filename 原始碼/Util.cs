using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace TenderDocGen
{
    /// <summary>共用小工具：字串正規化、資料夾名稱清理、民國日期。</summary>
    static class Util
    {
        /// <summary>NFKC 正規化＋去頭尾空白（含全形空白）。全形冒號/數字/英文一併轉半形。</summary>
        public static string Nfkc(string s)
        {
            if (s == null) return "";
            return s.Normalize(NormalizationForm.FormKC).Trim().Trim('　', ' ', '\t');
        }

        /// <summary>值是否代表「是」。空白視為「是」（預設產生），必須明確填否定詞才略過。</summary>
        public static bool IsTruthy(string s)
        {
            string v = Nfkc(s).ToUpperInvariant();
            if (v == "") return true;
            if (v == "否" || v == "N" || v == "NO" || v == "FALSE" || v == "0" ||
                v == "X" || v == "✗" || v == "不" || v == "不要" || v == "免") return false;
            return true;
        }

        static readonly Dictionary<char, char> FullWidthMap = new Dictionary<char, char>
        {
            { '<', '＜' }, { '>', '＞' }, { ':', '：' }, { '"', '＂' },
            { '/', '／' }, { '\\', '＼' }, { '|', '｜' }, { '?', '？' }, { '*', '＊' }
        };

        static readonly HashSet<string> ReservedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
        };

        /// <summary>
        /// 把標案名稱轉成安全的 Windows 資料夾名：
        /// 非法字元轉全形、去控制字元、去結尾點與空白、保留裝置名前綴底線、過長截斷。
        /// 回傳 null 表示清理後為空（呼叫端應報錯）。
        /// </summary>
        public static string SanitizeFolderName(string name, int maxLength, List<string> warnings)
        {
            if (name == null) return null;
            StringBuilder sb = new StringBuilder();
            foreach (char c in name.Trim())
            {
                if (char.IsControl(c)) continue;
                char mapped;
                if (FullWidthMap.TryGetValue(c, out mapped)) sb.Append(mapped);
                else sb.Append(c);
            }
            string result = sb.ToString().TrimEnd('.', ' ', '　');
            if (result.Length == 0) return null;
            if (ReservedNames.Contains(result))
            {
                result = "＿" + result;
                warnings.Add("標案名稱為 Windows 保留字，已加上前綴：" + result);
            }
            if (result.Length > maxLength)
            {
                result = result.Substring(0, maxLength).TrimEnd('.', ' ', '　');
                warnings.Add("標案名稱過長，資料夾名已截短為：" + result);
            }
            if (result.Length == 0) return null;   // 截斷後可能全被去光（例如整串點號）
            return result;
        }

        /// <summary>
        /// 解析民國日期字串。接受「115年1月1日」「115/1/1」「115.1.1」「115-1-1」，
        /// 以及 Excel 誤存的西元日期序列值（自動轉民國）。
        /// </summary>
        public static bool TryParseRocDate(string s, out int year, out int month, out int day)
        {
            year = 0; month = 0; day = 0;
            string v = Nfkc(s);
            if (v == "") return false;

            Match m = Regex.Match(v, @"^(\d{2,3})\s*[年/\.\-]\s*(\d{1,2})\s*[月/\.\-]\s*(\d{1,2})\s*日?$");
            if (m.Success)
            {
                year = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                month = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
                day = int.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture);
                // 用實際曆法驗證（含各月日數與閏年），拒絕 115年2月30日 等不存在的日期
                return IsValidRocDate(year, month, day);
            }

            // Excel 日期序列值（儲存格被格式化成日期時會拿到數字）
            double serial;
            if (double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out serial) &&
                serial > 20000 && serial < 80000)
            {
                DateTime dt = DateTime.FromOADate(serial);
                year = dt.Year - 1911; month = dt.Month; day = dt.Day;
                return true;
            }
            return false;
        }

        /// <summary>民國年月日是否為真實存在的日期（西元年 = 民國年 + 1911）。</summary>
        static bool IsValidRocDate(int rocYear, int month, int day)
        {
            if (month < 1 || month > 12 || day < 1) return false;
            int gYear = rocYear + 1911;
            if (gYear < 1 || gYear > 9999) return false;
            return day <= DateTime.DaysInMonth(gYear, month);
        }

        /// <summary>今天的民國日期。</summary>
        public static void TodayRoc(out int year, out int month, out int day)
        {
            DateTime now = DateTime.Now;
            year = now.Year - 1911; month = now.Month; day = now.Day;
        }

        /// <summary>把 Excel 可能存成科學記號的數字還原成完整數字字串（電話欄防呆）。</summary>
        public static string ExpandScientific(string raw)
        {
            if (raw == null) return "";
            if (raw.IndexOf('E') < 0 && raw.IndexOf('e') < 0) return raw;
            double d;
            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                return d.ToString("0.####################", CultureInfo.InvariantCulture);
            return raw;
        }
    }
}
