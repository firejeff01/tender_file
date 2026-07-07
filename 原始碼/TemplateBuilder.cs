using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace TenderDocGen
{
    /// <summary>從原始 ODT 擷取到的一段「彩色文字」，代表一個候選參數。</summary>
    class ColoredRun
    {
        public string Id;            // 穩定識別（R0、R1…），對應 Build 的 token map
        public string ColorHex;      // 例：#FF0000
        public string Text;          // 合併後的顯示文字
        public string Category;      // red / blue / other（決定新參數的預設類型）
        internal List<XElement> Spans = new List<XElement>();
    }

    /// <summary>
    /// 把「使用者自製的彩色 ODT」正規化成 ${參數} 範本。
    /// 沿用與 工具\normalize.ps1 相同的規則（相鄰同色 span 合併、PARAM.* 專屬樣式、
    /// 確定性重打包），但辨識參數改為「擷取彩色 run + 使用者對應」以支援任意新文件。
    /// </summary>
    class TemplateBuilder
    {
        static readonly XNamespace NsOffice = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        static readonly XNamespace NsText = OdtWriter.NsText;
        static readonly XNamespace NsStyle = OdtWriter.NsStyle;
        static readonly XNamespace NsFo = OdtWriter.NsFo;

        readonly List<KeyValuePair<string, byte[]>> _entries;
        readonly int _contentIdx;
        readonly XDocument _doc;
        readonly Dictionary<string, string> _colorOf = new Dictionary<string, string>(); // styleName -> colorHex(非黑)
        List<ColoredRun> _runs;

        TemplateBuilder(List<KeyValuePair<string, byte[]>> entries, int contentIdx, XDocument doc)
        {
            _entries = entries;
            _contentIdx = contentIdx;
            _doc = doc;
            BuildColorMap();
        }

        public static TemplateBuilder Load(string odtPath)
        {
            List<KeyValuePair<string, byte[]>> entries = OdtWriter.ReadAllEntries(odtPath);
            int idx = entries.FindIndex(e => e.Key == "content.xml");
            if (idx < 0)
                throw new InvalidDataException("這個檔案缺少 content.xml，不是有效的 ODT 文件。");
            XDocument doc;
            try
            {
                doc = XDocument.Parse(Encoding.UTF8.GetString(OdtWriter.StripBom(entries[idx].Value)),
                    LoadOptions.PreserveWhitespace);
            }
            catch (XmlException ex)
            {
                throw new InvalidDataException("這個 ODT 的 content.xml 無法解析：" + ex.Message);
            }
            return new TemplateBuilder(entries, idx, doc);
        }

        void BuildColorMap()
        {
            XElement auto = _doc.Descendants(NsOffice + "automatic-styles").FirstOrDefault();
            if (auto == null) return;
            foreach (XElement st in auto.Elements(NsStyle + "style"))
            {
                XAttribute nm = st.Attribute(NsStyle + "name");
                if (nm == null) continue;
                XElement tp = st.Element(NsStyle + "text-properties");
                if (tp == null) continue;
                XAttribute c = tp.Attribute(NsFo + "color");
                if (c == null) continue;
                string hex = c.Value.Trim().ToUpperInvariant();
                if (hex == "" || hex == "#000000") continue;   // 黑色不算參數候選
                _colorOf[nm.Value] = hex;
            }
        }

        /// <summary>擷取彩色 run（相鄰同色合併）。空白／純底線等無意義片段不列入。</summary>
        public List<ColoredRun> ExtractRuns()
        {
            if (_runs != null) return _runs;
            _runs = new List<ColoredRun>();

            XElement body = _doc.Descendants(NsOffice + "body").FirstOrDefault();
            if (body == null) return _runs;

            List<XElement> spans = body.Descendants(NsText + "span").ToList();
            HashSet<XElement> consumed = new HashSet<XElement>();
            int nextId = 0;

            foreach (XElement sp in spans)
            {
                if (consumed.Contains(sp)) continue;
                string color;
                if (!_colorOf.TryGetValue(StyleNameOf(sp), out color)) continue;

                // 從此 span 開始，往後合併「相鄰且同色」的 span
                List<XElement> group = new List<XElement> { sp };
                consumed.Add(sp);
                XElement cur = sp;
                while (true)
                {
                    XElement nxt = AdjacentSpan(cur);
                    if (nxt == null || consumed.Contains(nxt)) break;
                    string ncolor;
                    if (!_colorOf.TryGetValue(StyleNameOf(nxt), out ncolor) || ncolor != color) break;
                    group.Add(nxt);
                    consumed.Add(nxt);
                    cur = nxt;
                }

                string text = string.Concat(group.Select(g => g.Value));
                if (text.Trim(' ', '　', '\t', '\r', '\n', '_').Length == 0) continue; // 空白/底線 junk 不算參數

                ColoredRun run = new ColoredRun();
                run.Id = "R" + (nextId++);
                run.ColorHex = color;
                run.Text = text;
                run.Category = Classify(color);
                run.Spans = group;
                _runs.Add(run);
            }
            return _runs;
        }

        /// <summary>
        /// 依 run→token 對應表產生正規化後的 ODT bytes。
        /// map 中不含或值為空的 run 視為「略過」（保留原樣）。
        /// </summary>
        public byte[] Build(Dictionary<string, string> runToToken)
        {
            if (_runs == null) ExtractRuns();
            Dictionary<string, string> styleCache = new Dictionary<string, string>();

            foreach (ColoredRun run in _runs)
            {
                string token;
                if (!runToToken.TryGetValue(run.Id, out token) || string.IsNullOrEmpty(token)) continue;
                // 巢狀彩色 span 情況下，子 span 可能已隨父 span 一併被移除 → 跳過，避免 ReplaceWith 拋例外
                if (run.Spans.Count == 0 || run.Spans[0].Parent == null) continue;

                string baseStyle = StyleNameOf(run.Spans[0]);
                string paramStyle = EnsureParamStyle(styleCache, token, baseStyle);

                XElement newSpan = new XElement(NsText + "span",
                    new XAttribute(NsText + "style-name", paramStyle),
                    "${" + token + "}");
                // AdjacentSpan 合併時會穿透 span 之間「只有空白」的文字節點；這些空白不屬於任何一段，
                // 若不清掉會殘留在 ${token} 旁邊（多出空格）。先蒐集再一併移除。
                List<XText> tunneledWs = new List<XText>();
                for (int i = 0; i + 1 < run.Spans.Count; i++)
                {
                    XNode n = run.Spans[i].NextNode;
                    while (n != null && n != run.Spans[i + 1])
                    {
                        XText t = n as XText;
                        if (t != null) tunneledWs.Add(t);
                        n = n.NextNode;
                    }
                }
                run.Spans[0].ReplaceWith(newSpan);
                for (int i = 1; i < run.Spans.Count; i++)
                    if (run.Spans[i].Parent != null) run.Spans[i].Remove();
                foreach (XText ws in tunneledWs)
                    if (ws.Parent != null) ws.Remove();
            }

            RemoveOrphanColorStyles();

            byte[] contentBytes = SerializeXml(_doc);
            _entries[_contentIdx] = new KeyValuePair<string, byte[]>("content.xml", contentBytes);
            return OdtWriter.Repack(_entries);
        }

        // ---------- 內部 ----------

        static string StyleNameOf(XElement span)
        {
            XAttribute a = span.Attribute(NsText + "style-name");
            return a != null ? a.Value : "";
        }

        // 下一個「緊鄰的 sibling span」：可穿透 bookmark 標記與純空白文字，遇到有內容的文字或
        // 其他元素則中斷（與 normalize.ps1 的 Next-AdjacentSpan 一致）。
        static XElement AdjacentSpan(XElement span)
        {
            XNode n = span.NextNode;
            while (n != null)
            {
                XText t = n as XText;
                if (t != null)
                {
                    if (t.Value.Trim(' ', '　', '\t', '\r', '\n').Length > 0) return null;
                }
                else
                {
                    XElement e = n as XElement;
                    if (e != null)
                    {
                        string ln = e.Name.LocalName;
                        if (ln == "span") return e;
                        if (ln == "bookmark-start" || ln == "bookmark-end" || ln == "bookmark") { }
                        else return null;
                    }
                }
                n = n.NextNode;
            }
            return null;
        }

        // 建立（或重用）PARAM 專屬樣式：複製 base 自動樣式、強制紅色（範本中醒目標示參數）。
        string EnsureParamStyle(Dictionary<string, string> cache, string token, string baseStyle)
        {
            string safeBase = Regex.Replace(baseStyle ?? "", @"[^0-9A-Za-z一-鿿]", "_");
            string name = OdtWriter.ParamStylePrefix + token + "." + safeBase;
            if (cache.ContainsKey(name)) return name;

            XElement auto = _doc.Descendants(NsOffice + "automatic-styles").FirstOrDefault();
            if (auto == null)
            {
                // 極少見：沒有 automatic-styles 容器，補一個放在 body 之前
                auto = new XElement(NsOffice + "automatic-styles");
                XElement bodyEl = _doc.Descendants(NsOffice + "body").First();
                bodyEl.AddBeforeSelf(auto);
            }

            XElement baseEl = auto.Elements(NsStyle + "style")
                .FirstOrDefault(s => (string)s.Attribute(NsStyle + "name") == baseStyle);
            XElement clone;
            if (baseEl != null)
            {
                clone = new XElement(baseEl);
            }
            else
            {
                clone = new XElement(NsStyle + "style",
                    new XAttribute(NsStyle + "family", "text"),
                    new XElement(NsStyle + "text-properties"));
            }
            clone.SetAttributeValue(NsStyle + "name", name);
            XElement tp = clone.Element(NsStyle + "text-properties");
            if (tp == null) { tp = new XElement(NsStyle + "text-properties"); clone.Add(tp); }
            tp.SetAttributeValue(NsFo + "color", "#FF0000");
            auto.Add(clone);

            cache[name] = name;
            return name;
        }

        // 移除已無任何 span 參照的彩色自動樣式（被替換掉的 run 遺留的定義）。
        void RemoveOrphanColorStyles()
        {
            HashSet<string> used = new HashSet<string>();
            foreach (XElement sp in _doc.Descendants(NsText + "span"))
            {
                string sn = StyleNameOf(sp);
                if (sn != "") used.Add(sn);
            }
            XElement auto = _doc.Descendants(NsOffice + "automatic-styles").FirstOrDefault();
            if (auto == null) return;
            foreach (string styleName in _colorOf.Keys.ToList())
            {
                if (used.Contains(styleName)) continue;
                XElement st = auto.Elements(NsStyle + "style")
                    .FirstOrDefault(s => (string)s.Attribute(NsStyle + "name") == styleName);
                if (st != null) st.Remove();
            }
        }

        static string Classify(string hex)
        {
            int r, g, b;
            if (!TryParseHex(hex, out r, out g, out b)) return "other";
            if (r >= 0x80 && g < 0x80 && b < 0x80) return "red";
            if (b >= 0x80 && r < 0x80) return "blue";
            return "other";
        }

        static bool TryParseHex(string hex, out int r, out int g, out int b)
        {
            r = g = b = 0;
            if (hex == null) return false;
            hex = hex.TrimStart('#');
            if (hex.Length != 6) return false;
            return int.TryParse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out r)
                && int.TryParse(hex.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out g)
                && int.TryParse(hex.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out b);
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

        /// <summary>驗證正規化產出：zip mimetype 第一且 stored、content.xml well-formed、含 ${ token。</summary>
        public static List<string> ValidateTemplateBytes(byte[] bytes)
        {
            List<string> problems = new List<string>();
            if (bytes.Length < 40 || bytes[0] != 0x50 || bytes[1] != 0x4B)
            { problems.Add("不是有效的 zip"); return problems; }
            int method = bytes[8] | (bytes[9] << 8);
            int nameLen = bytes[26] | (bytes[27] << 8);
            string firstName = nameLen <= 30 ? Encoding.ASCII.GetString(bytes, 30, Math.Min(nameLen, bytes.Length - 30)) : "";
            if (firstName != "mimetype") problems.Add("第一個 entry 不是 mimetype");
            if (method != 0) problems.Add("mimetype 未以 stored 儲存");
            try
            {
                using (MemoryStream ms = new MemoryStream(bytes))
                using (System.IO.Compression.ZipArchive zip =
                    new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read))
                {
                    System.IO.Compression.ZipArchiveEntry ce = zip.GetEntry("content.xml");
                    if (ce == null) { problems.Add("缺少 content.xml"); return problems; }
                    using (Stream s = ce.Open()) XDocument.Load(s);   // well-formed 驗證
                }
            }
            catch (Exception ex) { problems.Add("content.xml 驗證失敗：" + ex.Message); }
            return problems;
        }
    }
}
