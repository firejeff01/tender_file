using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace TenderDocGen
{
    /// <summary>產出 ODT 的結構自我驗證（每個檔案產生後都會跑）。</summary>
    static class Validator
    {
        static readonly XNamespace NsOffice = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

        /// <summary>
        /// 回傳問題清單（空 = 通過）。
        /// knownTokens：此範本的參數名集合；用於精準偵測「未替換的 ${參數}」而不會誤判
        /// 使用者資料中恰好含有的 ${ 字面（例如標案名稱寫成 AI${x}平台）。
        /// </summary>
        public static List<string> ValidateOdt(string path, Dictionary<string, string> expectedValues,
                                               ICollection<string> knownTokens)
        {
            List<string> problems = new List<string>();
            byte[] head;
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    head = new byte[64];
                    int n = fs.Read(head, 0, head.Length);
                    if (n < 40) { problems.Add("檔案過小"); return problems; }
                }
            }
            catch (Exception ex) { problems.Add("無法讀取檔案：" + ex.Message); return problems; }

            // --- zip 結構：第一個 local header 必須是未壓縮的 mimetype ---
            if (head[0] != 0x50 || head[1] != 0x4B)
                problems.Add("不是 zip 檔（缺少 PK 簽章）");
            else
            {
                int method = head[8] | (head[9] << 8);
                int nameLen = head[26] | (head[27] << 8);
                string firstName = nameLen <= 30 ? Encoding.ASCII.GetString(head, 30, Math.Min(nameLen, head.Length - 30)) : "";
                if (firstName != "mimetype") problems.Add("第一個 entry 不是 mimetype（ODF 規範要求）");
                if (method != 0) problems.Add("mimetype 未以 stored 方式儲存");
            }

            // --- 內容驗證 ---
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    ZipArchiveEntry mime = zip.GetEntry("mimetype");
                    if (mime == null) problems.Add("缺少 mimetype entry");
                    else
                    {
                        using (StreamReader sr = new StreamReader(mime.Open(), Encoding.ASCII))
                        {
                            string mt = sr.ReadToEnd();
                            if (mt != "application/vnd.oasis.opendocument.text")
                                problems.Add("mimetype 內容不正確：" + mt);
                        }
                    }

                    ZipArchiveEntry content = zip.GetEntry("content.xml");
                    if (content == null) { problems.Add("缺少 content.xml"); return problems; }

                    XDocument xdoc;
                    using (Stream s = content.Open()) xdoc = XDocument.Load(s);   // well-formed 驗證

                    XElement body = xdoc.Descendants(NsOffice + "body").FirstOrDefault();
                    string bodyText = body != null ? body.Value : xdoc.Root.Value;

                    // 只把「已知參數名」的 ${參數} 視為未替換殘留，避免誤判使用者資料裡的 ${ 字面
                    if (knownTokens != null)
                    {
                        foreach (string tk in knownTokens)
                            if (bodyText.Contains("${" + tk + "}"))
                            { problems.Add("文件內仍殘留未替換的參數「" + tk + "」"); break; }
                    }
                    else if (bodyText.Contains("${"))
                    {
                        problems.Add("文件內仍殘留未替換的 ${ 參數");
                    }

                    if (expectedValues != null)
                    {
                        // 換行/tab/連續空白在 ODF 是元素不是文字，比對時把空白類字元全部剔除
                        string bodyCompact = StripWs(bodyText);
                        foreach (KeyValuePair<string, string> kv in expectedValues)
                        {
                            string firstLine = kv.Value.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')[0];
                            string expect = StripWs(firstLine);
                            if (expect != "" && !bodyCompact.Contains(expect))
                                problems.Add("參數「" + kv.Key + "」的值未出現在文件中");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                problems.Add("內容驗證失敗：" + ex.Message);
            }
            return problems;
        }

        static string StripWs(string s)
        {
            StringBuilder sb = new StringBuilder(s.Length);
            foreach (char c in s)
                if (c != ' ' && c != '\t' && c != '　' && c != '\r' && c != '\n') sb.Append(c);
            return sb.ToString();
        }
    }
}
