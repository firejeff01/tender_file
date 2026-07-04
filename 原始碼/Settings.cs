using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TenderDocGen
{
    /// <summary>
    /// 可攜式設定：存於 exe 旁的「設定.txt」（key=value，UTF-8）。
    /// 跟著資料夾走，免安裝、免登錄檔。刪掉即還原預設。
    /// </summary>
    class Settings
    {
        readonly string _path;
        readonly Dictionary<string, string> _map =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public const string KeyOutputDir = "輸出資料夾";

        public Settings(string baseDir)
        {
            _path = Path.Combine(baseDir, "設定.txt");
            Load();
        }

        void Load()
        {
            try
            {
                if (!File.Exists(_path)) return;
                foreach (string line in File.ReadAllLines(_path, Encoding.UTF8))
                {
                    string t = line.Trim();
                    if (t == "" || t.StartsWith("#")) continue;
                    int i = t.IndexOf('=');
                    if (i <= 0) continue;
                    _map[t.Substring(0, i).Trim()] = t.Substring(i + 1).Trim();
                }
            }
            catch { /* 設定讀取失敗就用預設，不影響運作 */ }
        }

        public string Get(string key, string fallback)
        {
            string v;
            if (_map.TryGetValue(key, out v) && v != "") return v;
            return fallback;
        }

        public void Set(string key, string value)
        {
            _map[key] = value;
            Save();
        }

        void Save()
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("# 標案文件產生器 設定檔（可整個刪除以還原預設值）");
                foreach (KeyValuePair<string, string> kv in _map)
                    sb.AppendLine(kv.Key + "=" + kv.Value);
                File.WriteAllText(_path, sb.ToString(), new UTF8Encoding(true));
            }
            catch { /* 無法寫入設定不影響本次運作 */ }
        }
    }
}
