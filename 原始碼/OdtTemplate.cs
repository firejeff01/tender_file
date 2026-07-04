using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace TenderDocGen
{
    /// <summary>一個 ODT 範本（含 tokens.txt 記錄的預期 token 次數，用於偵測範本被編輯過）。</summary>
    class TemplateInfo
    {
        public string FileName;                     // 例：保密切結書.odt
        public string BaseName;                     // 例：保密切結書
        public string FullPath;
        // tokens.txt 記錄的預期次數（僅供竄改預檢；缺 tokens.txt 時為空）
        public Dictionary<string, int> ExpectedTokenCounts = new Dictionary<string, int>();
        // 直接掃描範本內容得到的實際 token（產生所需參數的真正來源，不依賴 tokens.txt）
        public HashSet<string> Tokens = new HashSet<string>();
        public string LoadError;                    // 範本本身無法解析時的錯誤（null = 正常）
    }

    /// <summary>掃描範本資料夾＋讀取 tokens.txt manifest。</summary>
    class TemplateStore
    {
        public List<TemplateInfo> Templates = new List<TemplateInfo>();

        public static TemplateStore Load(string templateDir)
        {
            if (!Directory.Exists(templateDir))
                throw new DirectoryNotFoundException("找不到範本資料夾：" + templateDir);

            TemplateStore store = new TemplateStore();
            foreach (string path in Directory.GetFiles(templateDir, "*.odt").OrderBy(p => p))
            {
                string fileName = Path.GetFileName(path);
                if (fileName.StartsWith("~") || fileName.StartsWith(".")) continue; // 鎖定/暫存檔
                TemplateInfo info = new TemplateInfo();
                info.FileName = fileName;
                info.BaseName = Path.GetFileNameWithoutExtension(path);
                info.FullPath = path;
                // 直接從範本內容抓出所有 ${token}，作為「這份文件需要哪些參數」的權威來源。
                // 如此即使 tokens.txt 遺失或損壞，產生流程仍能正確判斷所需欄位。
                try
                {
                    foreach (string tk in OdtWriter.ExtractTokens(path)) info.Tokens.Add(tk);
                }
                catch (Exception ex) { info.LoadError = ex.Message; }
                store.Templates.Add(info);
            }
            if (store.Templates.Count == 0)
                throw new InvalidDataException("範本資料夾內沒有任何 .odt 範本：" + templateDir);

            // tokens.txt：每行「檔名|token|次數」
            string manifestPath = Path.Combine(templateDir, "tokens.txt");
            if (File.Exists(manifestPath))
            {
                foreach (string line in File.ReadAllLines(manifestPath, Encoding.UTF8))
                {
                    string trimmed = line.Trim();
                    if (trimmed == "" || trimmed.StartsWith("#")) continue;
                    string[] parts = trimmed.Split('|');
                    if (parts.Length != 3) continue;
                    TemplateInfo tpl = store.Templates.FirstOrDefault(t => t.FileName == parts[0]);
                    int count;
                    if (tpl != null && int.TryParse(parts[2], out count))
                        tpl.ExpectedTokenCounts[parts[1]] = count;
                }
            }
            return store;
        }

        /// <summary>所有範本用到的 token 聯集（給 Excel 欄位對應檢查用）。</summary>
        public HashSet<string> AllTokens()
        {
            HashSet<string> set = new HashSet<string>();
            foreach (TemplateInfo t in Templates)
                foreach (string k in t.ExpectedTokenCounts.Keys) set.Add(k);
            return set;
        }

        /// <summary>
        /// 掃描範本資料夾內所有 .odt，重新產生 tokens.txt（每行「檔名|token|次數」）。
        /// 新增／更新範本後呼叫，讓 manifest 與實際範本一致、避免漂移。
        /// </summary>
        public static void RewriteTokensFile(string templateDir)
        {
            List<string> lines = new List<string>();
            foreach (string path in Directory.GetFiles(templateDir, "*.odt").OrderBy(p => p))
            {
                string fileName = Path.GetFileName(path);
                if (fileName.StartsWith("~") || fileName.StartsWith(".")) continue;
                Dictionary<string, int> counts;
                try { counts = OdtWriter.ExtractTokenCounts(path); }
                catch { continue; }   // 無法解析的檔案跳過，不讓它擋住 manifest 重寫
                foreach (KeyValuePair<string, int> kv in counts.OrderBy(k => k.Key))
                    lines.Add(fileName + "|" + kv.Key + "|" + kv.Value);
            }
            File.WriteAllLines(Path.Combine(templateDir, "tokens.txt"), lines, new UTF8Encoding(true));
        }
    }

    /// <summary>ODT 參數替換與確定性重打包。</summary>
    static class OdtWriter
    {
        // 供 TemplateBuilder 等重用（正規化新範本時共用相同的命名空間與慣例）
        internal static readonly XNamespace NsText = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        internal static readonly XNamespace NsStyle = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        internal static readonly XNamespace NsFo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        internal const string ParamStylePrefix = "PARAM.";
        static readonly System.Text.RegularExpressions.Regex TokenRegex =
            new System.Text.RegularExpressions.Regex(@"^\$\{([^}]+)\}$");

        /// <summary>掃描範本 content.xml，回傳所有 PARAM span 內的 token 名稱（去重）。</summary>
        public static HashSet<string> ExtractTokens(string odtPath)
        {
            return new HashSet<string>(ExtractTokenCounts(odtPath).Keys);
        }

        /// <summary>掃描範本 content.xml，回傳每個 token 的出現次數（給 tokens.txt 與竄改預檢用）。</summary>
        public static Dictionary<string, int> ExtractTokenCounts(string odtPath)
        {
            Dictionary<string, int> counts = new Dictionary<string, int>();
            List<KeyValuePair<string, byte[]>> entries = ReadAllEntries(odtPath);
            int idx = entries.FindIndex(e => e.Key == "content.xml");
            if (idx < 0) throw new InvalidDataException("缺少 content.xml，不是有效的 ODT。");
            XDocument xdoc = XDocument.Parse(Encoding.UTF8.GetString(StripBom(entries[idx].Value)),
                LoadOptions.PreserveWhitespace);
            foreach (XElement span in xdoc.Descendants(NsText + "span"))
            {
                XAttribute sn = span.Attribute(NsText + "style-name");
                if (sn == null || !sn.Value.StartsWith(ParamStylePrefix)) continue;
                System.Text.RegularExpressions.Match m = TokenRegex.Match(span.Value);
                if (m.Success)
                {
                    string tk = m.Groups[1].Value;
                    counts[tk] = (counts.ContainsKey(tk) ? counts[tk] : 0) + 1;
                }
            }
            return counts;
        }

        /// <summary>
        /// 讀範本 → 預檢 → 替換 ${token} → 顏色處理 → 重打包成 byte[]。
        /// usedTokens 回傳此範本實際替換的參數（驗證器只驗這些）。
        /// 發現問題丟 InvalidDataException（訊息可直接顯示給使用者）。
        /// </summary>
        public static byte[] Generate(TemplateInfo tpl, Dictionary<string, string> values, bool keepRed,
                                      out HashSet<string> usedTokens)
        {
            usedTokens = new HashSet<string>();
            // --- 讀入範本全部 entry（保留原順序）---
            List<KeyValuePair<string, byte[]>> entries = ReadAllEntries(tpl.FullPath);
            int contentIdx = entries.FindIndex(e => e.Key == "content.xml");
            if (contentIdx < 0)
                throw new InvalidDataException("範本「" + tpl.FileName + "」缺少 content.xml，不是有效的 ODT。");

            string contentXml = Encoding.UTF8.GetString(StripBom(entries[contentIdx].Value));
            XDocument xdoc;
            try { xdoc = XDocument.Parse(contentXml, LoadOptions.PreserveWhitespace); }
            catch (XmlException ex)
            {
                throw new InvalidDataException("範本「" + tpl.FileName + "」的 content.xml 解析失敗：" + ex.Message);
            }

            // --- 預檢：PARAM span 的 token 統計必須符合 manifest（偵測範本被重存/編輯）---
            List<XElement> paramSpans = new List<XElement>();
            Dictionary<string, int> found = new Dictionary<string, int>();
            foreach (XElement span in xdoc.Descendants(NsText + "span"))
            {
                XAttribute sn = span.Attribute(NsText + "style-name");
                if (sn == null || !sn.Value.StartsWith(ParamStylePrefix)) continue;
                paramSpans.Add(span);
                Match m = Regex.Match(span.Value, @"^\$\{([^}]+)\}$");
                if (!m.Success)
                    throw new InvalidDataException(
                        "範本「" + tpl.FileName + "」的參數位置內容異常（可能被編輯過）：「" + span.Value + "」\n" +
                        "請取回原始正規化範本，勿以 Word/LibreOffice 重新儲存範本檔。");
                string token = m.Groups[1].Value;
                found[token] = (found.ContainsKey(token) ? found[token] : 0) + 1;
            }
            if (tpl.ExpectedTokenCounts.Count > 0)
            {
                foreach (KeyValuePair<string, int> kv in tpl.ExpectedTokenCounts)
                {
                    int actual = found.ContainsKey(kv.Key) ? found[kv.Key] : 0;
                    if (actual != kv.Value)
                        throw new InvalidDataException(
                            "範本「" + tpl.FileName + "」疑似被編輯過：參數「" + kv.Key + "」" +
                            "應出現 " + kv.Value + " 次，實際 " + actual + " 次。\n" +
                            "請取回原始正規化範本，勿以 Word/LibreOffice 重新儲存範本檔。");
                }
            }

            // --- 替換 ---
            foreach (XElement span in paramSpans)
            {
                string token = Regex.Match(span.Value, @"^\$\{([^}]+)\}$").Groups[1].Value;
                string val;
                if (!values.TryGetValue(token, out val) || val == null)
                    throw new InvalidDataException("缺少參數「" + token + "」的值（範本：" + tpl.FileName + "）。");
                usedTokens.Add(token);
                span.RemoveNodes();
                foreach (XNode n in BuildValueNodes(val)) span.Add(n);
            }

            // --- 顏色：預設把 PARAM.* 樣式改黑；設定「紅色」則保留紅字便於校對 ---
            if (!keepRed)
            {
                foreach (XElement st in xdoc.Descendants(NsStyle + "style"))
                {
                    XAttribute nm = st.Attribute(NsStyle + "name");
                    if (nm == null || !nm.Value.StartsWith(ParamStylePrefix)) continue;
                    XElement tp = st.Element(NsStyle + "text-properties");
                    if (tp != null) tp.SetAttributeValue(NsFo + "color", "#000000");
                }
            }

            // --- 序列化（UTF-8 無 BOM、不重排版）---
            entries[contentIdx] = new KeyValuePair<string, byte[]>("content.xml", SerializeXml(xdoc));

            // --- 確定性重打包 ---
            return Repack(entries);
        }

        /// <summary>把使用者輸入值轉成 ODF 節點：換行/定位字元/連續空白轉專用元素，其餘為文字（序列化時自動跳脫）。</summary>
        public static List<XNode> BuildValueNodes(string value)
        {
            string v = value.Replace("\r\n", "\n").Replace('\r', '\n');
            // 剔除 XML 1.0 非法控制字元（保留 \n \t）
            StringBuilder clean = new StringBuilder();
            foreach (char c in v)
                if (c == '\n' || c == '\t' || c >= 0x20) clean.Append(c);
            v = clean.ToString();

            List<XNode> nodes = new List<XNode>();
            StringBuilder text = new StringBuilder();
            int i = 0;
            while (i < v.Length)
            {
                char c = v[i];
                if (c == '\n')
                {
                    FlushText(nodes, text);
                    nodes.Add(new XElement(NsText + "line-break"));
                    i++;
                }
                else if (c == '\t')
                {
                    FlushText(nodes, text);
                    nodes.Add(new XElement(NsText + "tab"));
                    i++;
                }
                else if (c == ' ')
                {
                    int run = 0;
                    while (i + run < v.Length && v[i + run] == ' ') run++;
                    text.Append(' ');                       // 第一個空白留在文字裡
                    if (run > 1)
                    {
                        FlushText(nodes, text);
                        nodes.Add(new XElement(NsText + "s", new XAttribute(NsText + "c", run - 1)));
                    }
                    i += run;
                }
                else { text.Append(c); i++; }
            }
            FlushText(nodes, text);
            return nodes;
        }

        static void FlushText(List<XNode> nodes, StringBuilder text)
        {
            if (text.Length > 0) { nodes.Add(new XText(text.ToString())); text.Clear(); }
        }

        static byte[] SerializeXml(XDocument xdoc)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Encoding = new UTF8Encoding(false);
                settings.Indent = false;
                using (XmlWriter xw = XmlWriter.Create(ms, settings)) { xdoc.Save(xw); }
                return ms.ToArray();
            }
        }

        public static List<KeyValuePair<string, byte[]>> ReadAllEntries(string odtPath)
        {
            List<KeyValuePair<string, byte[]>> entries = new List<KeyValuePair<string, byte[]>>();
            using (FileStream fs = new FileStream(odtPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry e in zip.Entries)
                {
                    using (MemoryStream ms = new MemoryStream())
                    using (Stream s = e.Open())
                    {
                        s.CopyTo(ms);
                        entries.Add(new KeyValuePair<string, byte[]>(e.FullName, ms.ToArray()));
                    }
                }
            }
            return entries;
        }

        /// <summary>
        /// 重打包 ODT：mimetype 必為第一個 entry 且不壓縮（stored、無 extra field），
        /// 其餘依原順序 deflate。entry 名稱一律正斜線。
        /// 不用 .NET Framework 的 ZipArchive 寫入 —— 它的 NoCompression 仍以 Deflate
        /// method 寫入，違反 ODF 規範（國發會 ODF 檢測工具會驗 mimetype 為 stored）。
        /// 改以最小 zip writer 手工輸出 local header / central directory / EOCD。
        /// </summary>
        public static byte[] Repack(List<KeyValuePair<string, byte[]>> entries)
        {
            List<KeyValuePair<string, byte[]>> ordered = new List<KeyValuePair<string, byte[]>>();
            KeyValuePair<string, byte[]> mime = entries.FirstOrDefault(e => e.Key == "mimetype");
            if (mime.Key != null) ordered.Add(mime);
            foreach (KeyValuePair<string, byte[]> e in entries)
                if (e.Key != "mimetype")
                    ordered.Add(new KeyValuePair<string, byte[]>(e.Key.Replace('\\', '/'), e.Value));

            using (MemoryStream ms = new MemoryStream())
            {
                List<CentralRecord> central = new List<CentralRecord>();
                ushort dosTime, dosDate;
                ToDosDateTime(DateTime.Now, out dosDate, out dosTime);

                foreach (KeyValuePair<string, byte[]> e in ordered)
                {
                    bool stored = e.Key == "mimetype";
                    byte[] nameBytes = Encoding.UTF8.GetBytes(e.Key);
                    bool asciiName = e.Key.All(c => c < 128);
                    ushort flags = asciiName ? (ushort)0 : (ushort)0x0800;   // bit11 = UTF-8 名稱
                    byte[] data = stored ? e.Value : DeflateRaw(e.Value);
                    // deflate 反而變大時退回 stored（極小檔案可能發生）
                    ushort method = stored ? (ushort)0 : (ushort)8;
                    if (!stored && data.Length >= e.Value.Length) { data = e.Value; method = 0; }
                    uint crc = Crc32(e.Value);

                    CentralRecord rec = new CentralRecord();
                    rec.NameBytes = nameBytes;
                    rec.Flags = flags;
                    rec.Method = method;
                    rec.DosTime = dosTime;
                    rec.DosDate = dosDate;
                    rec.Crc = crc;
                    rec.CompSize = (uint)data.Length;
                    rec.UncompSize = (uint)e.Value.Length;
                    rec.LocalHeaderOffset = (uint)ms.Position;
                    central.Add(rec);

                    // local file header
                    WriteU32(ms, 0x04034b50);
                    WriteU16(ms, 20);                 // version needed
                    WriteU16(ms, flags);
                    WriteU16(ms, method);
                    WriteU16(ms, dosTime);
                    WriteU16(ms, dosDate);
                    WriteU32(ms, crc);
                    WriteU32(ms, rec.CompSize);
                    WriteU32(ms, rec.UncompSize);
                    WriteU16(ms, (ushort)nameBytes.Length);
                    WriteU16(ms, 0);                  // extra field 長度 = 0
                    ms.Write(nameBytes, 0, nameBytes.Length);
                    ms.Write(data, 0, data.Length);
                }

                // central directory
                long cdStart = ms.Position;
                foreach (CentralRecord rec in central)
                {
                    WriteU32(ms, 0x02014b50);
                    WriteU16(ms, 20);                 // version made by
                    WriteU16(ms, 20);                 // version needed
                    WriteU16(ms, rec.Flags);
                    WriteU16(ms, rec.Method);
                    WriteU16(ms, rec.DosTime);
                    WriteU16(ms, rec.DosDate);
                    WriteU32(ms, rec.Crc);
                    WriteU32(ms, rec.CompSize);
                    WriteU32(ms, rec.UncompSize);
                    WriteU16(ms, (ushort)rec.NameBytes.Length);
                    WriteU16(ms, 0);                  // extra
                    WriteU16(ms, 0);                  // comment
                    WriteU16(ms, 0);                  // disk number
                    WriteU16(ms, 0);                  // internal attrs
                    WriteU32(ms, 0);                  // external attrs
                    WriteU32(ms, rec.LocalHeaderOffset);
                    ms.Write(rec.NameBytes, 0, rec.NameBytes.Length);
                }
                long cdSize = ms.Position - cdStart;

                // end of central directory
                WriteU32(ms, 0x06054b50);
                WriteU16(ms, 0);
                WriteU16(ms, 0);
                WriteU16(ms, (ushort)central.Count);
                WriteU16(ms, (ushort)central.Count);
                WriteU32(ms, (uint)cdSize);
                WriteU32(ms, (uint)cdStart);
                WriteU16(ms, 0);

                return ms.ToArray();
            }
        }

        class CentralRecord
        {
            public byte[] NameBytes;
            public ushort Flags, Method, DosTime, DosDate;
            public uint Crc, CompSize, UncompSize, LocalHeaderOffset;
        }

        static byte[] DeflateRaw(byte[] data)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (DeflateStream ds = new DeflateStream(ms, CompressionLevel.Optimal, true))
                    ds.Write(data, 0, data.Length);
                return ms.ToArray();
            }
        }

        static void WriteU16(Stream s, ushort v) { s.WriteByte((byte)(v & 0xFF)); s.WriteByte((byte)(v >> 8)); }
        static void WriteU32(Stream s, uint v)
        {
            s.WriteByte((byte)(v & 0xFF)); s.WriteByte((byte)((v >> 8) & 0xFF));
            s.WriteByte((byte)((v >> 16) & 0xFF)); s.WriteByte((byte)((v >> 24) & 0xFF));
        }

        static void ToDosDateTime(DateTime dt, out ushort dosDate, out ushort dosTime)
        {
            if (dt.Year < 1980) dt = new DateTime(1980, 1, 1);
            dosDate = (ushort)(((dt.Year - 1980) << 9) | (dt.Month << 5) | dt.Day);
            dosTime = (ushort)((dt.Hour << 11) | (dt.Minute << 5) | (dt.Second / 2));
        }

        static uint[] _crcTable;
        static uint Crc32(byte[] data)
        {
            if (_crcTable == null)
            {
                uint[] table = new uint[256];
                for (uint i = 0; i < 256; i++)
                {
                    uint c = i;
                    for (int k = 0; k < 8; k++)
                        c = (c & 1) != 0 ? 0xEDB88320 ^ (c >> 1) : c >> 1;
                    table[i] = c;
                }
                _crcTable = table;
            }
            uint crc = 0xFFFFFFFF;
            foreach (byte b in data)
                crc = _crcTable[(crc ^ b) & 0xFF] ^ (crc >> 8);
            return crc ^ 0xFFFFFFFF;
        }

        internal static byte[] StripBom(byte[] bytes)
        {
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            {
                byte[] result = new byte[bytes.Length - 3];
                Array.Copy(bytes, 3, result, 0, result.Length);
                return result;
            }
            return bytes;
        }
    }
}
