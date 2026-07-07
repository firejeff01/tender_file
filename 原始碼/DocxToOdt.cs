using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace TenderDocGen
{
    /// <summary>
    /// 把一份 Word（.docx / OOXML）轉成結構完整、可通過 ODF 檢測的 .odt（記憶體 byte[]）。
    ///
    /// 用途：讓「新增範本」除了 ODT 也能吃 Word 檔——轉成 ODT 後，交給既有的
    /// <see cref="TemplateBuilder"/> 走完全相同的「彩色 run → ${參數}」正規化流程；
    /// 產出的範本與最終文件仍是 ODT（維持零 Office 依賴與 ODF 送件合規）。
    ///
    /// 支援範圍（依實際範本需求刻意收斂）：段落、文字格式（粗體/斜體/底線/刪除線/
    /// 顏色/字型/字級）、換行、Tab、對齊、縮排、行距、項目清單與編號、頁面大小與邊界。
    /// 碰到無法忠實轉換的結構（表格、圖片、文字方塊、OLE 物件、註腳、非空頁首/頁尾）
    /// 會丟 <see cref="InvalidDataException"/>（訊息可直接顯示給使用者），寧可報錯也不默默失真。
    /// 已知限制：硬分頁/分欄符會退化成一般換行；命名段落樣式的外觀（非清單編號）不繼承。
    /// </summary>
    static class DocxToOdt
    {
        // ---- OOXML（來源）命名空間 ----
        static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        static readonly XNamespace RelPkg = "http://schemas.openxmlformats.org/package/2006/relationships";

        // ---- ODF（產出）命名空間 ----
        static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        static readonly XNamespace Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        static readonly XNamespace Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        static readonly XNamespace Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        static readonly XNamespace Svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
        static readonly XNamespace Meta = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
        static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
        static readonly XNamespace Config = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
        static readonly XNamespace Manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

        const string DefaultWestern = "Times New Roman";
        const string DefaultAsian = "新細明體";

        public static byte[] Convert(string docxPath, out List<string> warnings)
        {
            byte[] bytes = File.ReadAllBytes(docxPath);
            return ConvertBytes(bytes, out warnings);
        }

        /// <summary>核心轉換（可被 selftest 直接以 byte[] 呼叫）。</summary>
        public static byte[] ConvertBytes(byte[] docx, out List<string> warnings)
        {
            warnings = new List<string>();

            Dictionary<string, byte[]> parts = ReadParts(docx);
            byte[] docBytes;
            if (!parts.TryGetValue("word/document.xml", out docBytes))
                throw new InvalidDataException("這不是有效的 Word 文件（缺少 word/document.xml）。");

            XDocument docXml = ParseXml(docBytes);
            XElement body = docXml.Root != null ? docXml.Root.Element(W + "body") : null;
            if (body == null)
                throw new InvalidDataException("Word 文件內容異常（找不到 <w:body>）。");

            RejectUnsupported(body, parts, docXml);

            // 讀清單定義（numId → 各層 numFmt / lvlText）
            NumberingDefs numbering = NumberingDefs.Load(parts);

            Converter conv = new Converter(numbering, warnings);
            XDocument content = conv.BuildContent(body);
            XDocument styles = conv.BuildStyles(body);

            // ---- 組成 ODT 各 entry ----
            List<KeyValuePair<string, byte[]>> entries = new List<KeyValuePair<string, byte[]>>();
            entries.Add(Entry("mimetype", Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.text")));
            entries.Add(Entry("content.xml", Serialize(content)));
            entries.Add(Entry("styles.xml", Serialize(styles)));
            entries.Add(Entry("meta.xml", Serialize(BuildMeta())));
            entries.Add(Entry("settings.xml", Serialize(BuildSettings())));
            entries.Add(Entry("META-INF/manifest.xml", Serialize(BuildManifest())));

            return OdtWriter.Repack(entries);
        }

        // =================================================================
        //  不支援結構偵測（寧可報錯也不失真）
        // =================================================================
        static void RejectUnsupported(XElement body, Dictionary<string, byte[]> parts, XDocument docXml)
        {
            List<string> found = new List<string>();
            AddIf(found, body.Descendants(W + "tbl").Any(), "表格");
            AddIf(found, body.Descendants(W + "drawing").Any() || body.Descendants(W + "pict").Any(), "圖片／繪圖物件");
            AddIf(found, body.Descendants(W + "object").Any(), "內嵌物件（OLE）");
            AddIf(found, body.Descendants(W + "txbxContent").Any(), "文字方塊");
            AddIf(found, body.Descendants(W + "footnoteReference").Any(), "註腳");
            AddIf(found, body.Descendants(W + "endnoteReference").Any(), "章節附註");

            // 頁首／頁尾：只有在被參照的頁首/頁尾實際含有文字時才擋（避免誤擋空白預設頁首/頁尾）
            CheckHeaderFooter(parts, docXml, found);

            if (found.Count > 0)
                throw new InvalidDataException(
                    "這份 Word 含有目前尚未支援的內容：" + string.Join("、", found.Distinct()) + "。\n" +
                    "請改用不含這些元素的版本，或先另存成 ODT 再新增。");
        }

        static void CheckHeaderFooter(Dictionary<string, byte[]> parts, XDocument docXml, List<string> found)
        {
            byte[] relBytes;
            if (!parts.TryGetValue("word/_rels/document.xml.rels", out relBytes)) return;
            XDocument rels;
            try { rels = ParseXml(relBytes); } catch { return; }

            bool header = false, footer = false;
            foreach (XElement r in rels.Descendants(RelPkg + "Relationship"))
            {
                string type = (string)r.Attribute("Type") ?? "";
                string target = (string)r.Attribute("Target") ?? "";
                if (target == "") continue;
                bool isHeader = type.EndsWith("/header");
                bool isFooter = type.EndsWith("/footer");
                if (!isHeader && !isFooter) continue;

                string partName = "word/" + target.TrimStart('/');
                byte[] hb;
                if (!parts.TryGetValue(partName, out hb)) continue;
                bool hasContent;
                try
                {
                    XDocument hd = ParseXml(hb);
                    string txt = string.Concat(hd.Descendants(W + "t").Select(t => t.Value));
                    // 不只看文字：只含 Logo/印信圖、表格、文字方塊的頁首/頁尾也算「非空」，須擋下不可默默丟掉
                    hasContent = txt.Trim().Length > 0
                        || hd.Descendants(W + "drawing").Any() || hd.Descendants(W + "pict").Any()
                        || hd.Descendants(W + "tbl").Any() || hd.Descendants(W + "txbxContent").Any()
                        || hd.Descendants(W + "object").Any();
                }
                catch { continue; }
                if (!hasContent) continue;   // 完全空白的頁首/頁尾不算
                if (isHeader) header = true;
                if (isFooter) footer = true;
            }
            AddIf(found, header, "頁首");
            AddIf(found, footer, "頁尾");
        }

        static void AddIf(List<string> list, bool cond, string label) { if (cond) list.Add(label); }

        // =================================================================
        //  轉換器本體
        // =================================================================
        class Converter
        {
            readonly NumberingDefs _num;
            readonly List<string> _warnings;

            // 產出的 content 自動樣式（去重）
            readonly Dictionary<string, string> _textStyleKey = new Dictionary<string, string>();
            readonly List<XElement> _textStyles = new List<XElement>();
            readonly Dictionary<string, string> _paraStyleKey = new Dictionary<string, string>();
            readonly List<XElement> _paraStyles = new List<XElement>();
            readonly Dictionary<string, string> _listStyleByNumId = new Dictionary<string, string>();
            readonly List<XElement> _listStyles = new List<XElement>();
            readonly HashSet<string> _fonts = new HashSet<string>();
            readonly HashSet<string> _warned = new HashSet<string>();   // 已提醒過的編號格式（避免重複）
            int _tSeq, _pSeq, _lSeq;

            public Converter(NumberingDefs num, List<string> warnings) { _num = num; _warnings = warnings; }

            public XDocument BuildContent(XElement body)
            {
                XElement text = new XElement(Office + "text");
                EmitBlocks(body, text);

                XElement autoStyles = new XElement(Office + "automatic-styles");
                foreach (XElement s in _paraStyles) autoStyles.Add(s);
                foreach (XElement s in _textStyles) autoStyles.Add(s);
                foreach (XElement s in _listStyles) autoStyles.Add(s);

                XElement root = new XElement(Office + "document-content",
                    new XAttribute(XNamespace.Xmlns + "office", Office.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "style", Style.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "text", Text.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "fo", Fo.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "svg", Svg.NamespaceName),
                    new XAttribute(Office + "version", "1.2"),
                    BuildFontDecls(),
                    autoStyles,
                    new XElement(Office + "body", text));
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            // 逐段落 / 清單輸出。連續的清單段落（同 numId）合成一個 text:list。
            void EmitBlocks(XElement body, XElement target)
            {
                List<XElement> paras = CollectBlockParas(body);
                int i = 0;
                while (i < paras.Count)
                {
                    XElement p = paras[i];
                    string numId = ListNumId(p);
                    if (numId == null)
                    {
                        target.Add(BuildParagraph(p));
                        i++;
                        continue;
                    }
                    // 收集連續、同 numId 的清單段落
                    List<XElement> groupItems = new List<XElement>();
                    while (i < paras.Count && ListNumId(paras[i]) == numId)
                    {
                        groupItems.Add(paras[i]);
                        i++;
                    }
                    target.Add(BuildList(numId, groupItems));
                }
            }

            // 逐層展開 body 內的段落：穿透區塊層級的 sdt（內容控制項），依序取出所有 w:p，
            // 避免被 sdt 包住的整段內容漏掉（表格等仍由 RejectUnsupported 事先擋下）。
            static List<XElement> CollectBlockParas(XElement container)
            {
                List<XElement> paras = new List<XElement>();
                foreach (XElement child in container.Elements())
                {
                    string ln = child.Name.LocalName;
                    if (ln == "p") paras.Add(child);
                    else if (ln == "sdt")
                    {
                        XElement c = child.Element(W + "sdtContent");
                        if (c != null) paras.AddRange(CollectBlockParas(c));
                    }
                }
                return paras;
            }

            // ---- 段落 ----
            XElement BuildParagraph(XElement p)
            {
                string styleName = EnsureParaStyle(p.Element(W + "pPr"));
                XElement outP = new XElement(Text + "p");
                if (styleName != null) outP.SetAttributeValue(Text + "style-name", styleName);
                EmitRuns(p, outP);
                return outP;
            }

            void EmitRuns(XElement p, XElement outP)
            {
                List<XElement> runs = new List<XElement>();
                CollectRuns(p, runs);
                foreach (XElement r in runs) EmitRun(r, outP);
            }

            // 逐層展開段落內的 run：穿透超連結、追蹤修訂插入(ins)、智慧標籤、內容控制項(sdt)、
            // 簡單功能變數(fldSimple)——包含彼此巢狀的情形——依文件順序收集所有 w:r。
            // 不進入 w:del（已刪除的追蹤修訂內容不應出現）。
            static void CollectRuns(XElement container, List<XElement> outRuns)
            {
                foreach (XElement child in container.Elements())
                {
                    string ln = child.Name.LocalName;
                    if (ln == "r") outRuns.Add(child);
                    else if (ln == "hyperlink" || ln == "smartTag" || ln == "ins" || ln == "fldSimple")
                        CollectRuns(child, outRuns);
                    else if (ln == "sdt")
                    {
                        XElement c = child.Element(W + "sdtContent");
                        if (c != null) CollectRuns(c, outRuns);
                    }
                }
            }

            void EmitRun(XElement r, XElement outP)
            {
                XElement rPr = r.Element(W + "rPr");
                string styleName = EnsureTextStyle(rPr);
                List<XNode> nodes = new List<XNode>();
                foreach (XElement child in r.Elements())
                {
                    string ln = child.Name.LocalName;
                    if (ln == "t")
                        nodes.AddRange(OdtWriter.BuildValueNodes(child.Value));
                    else if (ln == "tab")
                        nodes.Add(new XElement(Text + "tab"));
                    else if (ln == "br" || ln == "cr")
                        nodes.Add(new XElement(Text + "line-break"));
                }
                if (nodes.Count == 0) return;

                if (styleName == null)
                {
                    foreach (XNode n in nodes) outP.Add(n);   // 無格式 → 直接放進段落
                }
                else
                {
                    XElement span = new XElement(Text + "span", new XAttribute(Text + "style-name", styleName));
                    foreach (XNode n in nodes) span.Add(n);
                    outP.Add(span);
                }
            }

            // ---- 清單 ----
            XElement BuildList(string numId, List<XElement> items)
            {
                string listStyle = EnsureListStyle(numId);
                XElement rootList = new XElement(Text + "list", new XAttribute(Text + "style-name", listStyle));

                // 以 ilvl 建巢狀。levelStack[k] = 該層目前的 text:list 元素。
                List<XElement> levelLists = new List<XElement> { rootList };
                List<XElement> lastItemAtLevel = new List<XElement>();

                foreach (XElement p in items)
                {
                    int lvl = ListLevel(p);
                    if (lvl < 0) lvl = 0;
                    // 需要往下鑽：在上一層最後一個 list-item 底下建新的 text:list
                    while (levelLists.Count <= lvl)
                    {
                        int parentLvl = levelLists.Count - 1;
                        XElement parentItem = (lastItemAtLevel.Count > parentLvl && lastItemAtLevel[parentLvl] != null)
                            ? lastItemAtLevel[parentLvl]
                            : null;
                        XElement nested = new XElement(Text + "list");
                        if (parentItem != null) parentItem.Add(nested);
                        else levelLists[parentLvl].Add(new XElement(Text + "list-item", nested)); // 保底
                        levelLists.Add(nested);
                    }
                    // 往上回：丟掉比目前層更深的
                    while (levelLists.Count > lvl + 1) levelLists.RemoveAt(levelLists.Count - 1);
                    while (lastItemAtLevel.Count > lvl) lastItemAtLevel.RemoveAt(lastItemAtLevel.Count - 1);

                    XElement li = new XElement(Text + "list-item");
                    XElement para = BuildParagraph(p);
                    li.Add(para);
                    levelLists[lvl].Add(li);
                    while (lastItemAtLevel.Count <= lvl) lastItemAtLevel.Add(null);
                    lastItemAtLevel[lvl] = li;
                }
                return rootList;
            }

            // ================= 自動樣式建立（去重） =================

            string EnsureTextStyle(XElement rPr)
            {
                if (rPr == null) return null;
                List<XAttribute> props = new List<XAttribute>();

                if (BoolProp(rPr, "b"))
                {
                    props.Add(new XAttribute(Fo + "font-weight", "bold"));
                    props.Add(new XAttribute(Style + "font-weight-asian", "bold"));
                    props.Add(new XAttribute(Style + "font-weight-complex", "bold"));
                }
                if (BoolProp(rPr, "i"))
                {
                    props.Add(new XAttribute(Fo + "font-style", "italic"));
                    props.Add(new XAttribute(Style + "font-style-asian", "italic"));
                    props.Add(new XAttribute(Style + "font-style-complex", "italic"));
                }
                XElement u = rPr.Element(W + "u");
                if (u != null && ((string)u.Attribute(W + "val") ?? "single") != "none")
                {
                    string uval = (string)u.Attribute(W + "val") ?? "single";
                    props.Add(new XAttribute(Style + "text-underline-style", "solid"));
                    props.Add(new XAttribute(Style + "text-underline-width", "auto"));
                    props.Add(new XAttribute(Style + "text-underline-color", "font-color"));
                    if (uval == "double")
                        props.Add(new XAttribute(Style + "text-underline-type", "double"));
                }
                if (rPr.Element(W + "strike") != null && BoolProp(rPr, "strike"))
                {
                    props.Add(new XAttribute(Style + "text-line-through-style", "solid"));
                }
                // 顏色（參數偵測的關鍵：非 auto/非黑 → fo:color）
                XElement color = rPr.Element(W + "color");
                if (color != null)
                {
                    string cv = ((string)color.Attribute(W + "val") ?? "").Trim();
                    if (cv != "" && !cv.Equals("auto", StringComparison.OrdinalIgnoreCase) && IsHex6(cv))
                        props.Add(new XAttribute(Fo + "color", "#" + cv.ToUpperInvariant()));
                }
                // 字級（half-point → pt）
                XElement sz = rPr.Element(W + "sz");
                if (sz != null)
                {
                    string pt = HalfPointToPt((string)sz.Attribute(W + "val"));
                    if (pt != null)
                    {
                        props.Add(new XAttribute(Fo + "font-size", pt));
                        props.Add(new XAttribute(Style + "font-size-asian", pt));
                        props.Add(new XAttribute(Style + "font-size-complex", pt));
                    }
                }
                // 字型
                XElement rFonts = rPr.Element(W + "rFonts");
                if (rFonts != null)
                {
                    string ascii = (string)rFonts.Attribute(W + "ascii");
                    string ea = (string)rFonts.Attribute(W + "eastAsia");
                    string cs = (string)rFonts.Attribute(W + "cs");
                    if (!string.IsNullOrEmpty(ascii)) { props.Add(new XAttribute(Style + "font-name", ascii)); _fonts.Add(ascii); }
                    if (!string.IsNullOrEmpty(ea)) { props.Add(new XAttribute(Style + "font-name-asian", ea)); _fonts.Add(ea); }
                    if (!string.IsNullOrEmpty(cs)) { props.Add(new XAttribute(Style + "font-name-complex", cs)); _fonts.Add(cs); }
                }

                if (props.Count == 0) return null;

                XElement tp = new XElement(Style + "text-properties");
                foreach (XAttribute a in props) tp.Add(a);
                string key = tp.ToString(SaveOptions.DisableFormatting);
                string name;
                if (_textStyleKey.TryGetValue(key, out name)) return name;

                name = "T" + (++_tSeq);
                XElement st = new XElement(Style + "style",
                    new XAttribute(Style + "name", name),
                    new XAttribute(Style + "family", "text"),
                    tp);
                _textStyles.Add(st);
                _textStyleKey[key] = name;
                return name;
            }

            string EnsureParaStyle(XElement pPr)
            {
                if (pPr == null) return "Standard";
                List<XAttribute> props = new List<XAttribute>();

                XElement jc = pPr.Element(W + "jc");
                if (jc != null)
                {
                    string a = MapAlign((string)jc.Attribute(W + "val"));
                    if (a != null) props.Add(new XAttribute(Fo + "text-align", a));
                }
                XElement ind = pPr.Element(W + "ind");
                if (ind != null)
                {
                    string left = TwipsAttr(ind, "left", "start");
                    string right = TwipsAttr(ind, "right", "end");
                    if (left != null) props.Add(new XAttribute(Fo + "margin-left", left));
                    if (right != null) props.Add(new XAttribute(Fo + "margin-right", right));
                    string firstLine = (string)ind.Attribute(W + "firstLine");
                    string hanging = (string)ind.Attribute(W + "hanging");
                    if (!string.IsNullOrEmpty(hanging)) props.Add(new XAttribute(Fo + "text-indent", "-" + TwipsToCm(hanging)));
                    else if (!string.IsNullOrEmpty(firstLine)) props.Add(new XAttribute(Fo + "text-indent", TwipsToCm(firstLine)));
                }
                XElement spacing = pPr.Element(W + "spacing");
                if (spacing != null)
                {
                    string before = (string)spacing.Attribute(W + "before");
                    string after = (string)spacing.Attribute(W + "after");
                    if (!string.IsNullOrEmpty(before)) props.Add(new XAttribute(Fo + "margin-top", TwipsToCm(before)));
                    if (!string.IsNullOrEmpty(after)) props.Add(new XAttribute(Fo + "margin-bottom", TwipsToCm(after)));
                    string lineRule = (string)spacing.Attribute(W + "lineRule");
                    string line = (string)spacing.Attribute(W + "line");
                    if (!string.IsNullOrEmpty(line) && (lineRule == null || lineRule == "auto"))
                    {
                        int lv;
                        if (int.TryParse(line, NumberStyles.Integer, CultureInfo.InvariantCulture, out lv) && lv > 0)
                            props.Add(new XAttribute(Fo + "line-height",
                                Math.Round(lv / 240.0 * 100.0).ToString("0", CultureInfo.InvariantCulture) + "%"));
                    }
                }

                if (props.Count == 0) return "Standard";

                XElement pp = new XElement(Style + "paragraph-properties");
                foreach (XAttribute a in props) pp.Add(a);
                string key = pp.ToString(SaveOptions.DisableFormatting);
                string name;
                if (_paraStyleKey.TryGetValue(key, out name)) return name;

                name = "P" + (++_pSeq);
                XElement st = new XElement(Style + "style",
                    new XAttribute(Style + "name", name),
                    new XAttribute(Style + "family", "paragraph"),
                    new XAttribute(Style + "parent-style-name", "Standard"),
                    pp);
                _paraStyles.Add(st);
                _paraStyleKey[key] = name;
                return name;
            }

            string EnsureListStyle(string numId)
            {
                string name;
                if (_listStyleByNumId.TryGetValue(numId, out name)) return name;

                name = "L" + (++_lSeq);
                XElement listStyle = new XElement(Text + "list-style", new XAttribute(Style + "name", name));
                NumberingDefs.Abstract abs = _num.ForNumId(numId);
                for (int lvl = 0; lvl < 9; lvl++)
                {
                    NumberingDefs.Level def = abs != null ? abs.Level(lvl) : null;
                    string fmt = def != null ? def.NumFmt : "bullet";
                    string lvlText = def != null ? def.LvlText : "";
                    listStyle.Add(BuildLevelStyle(lvl, fmt, lvlText));
                }
                _listStyles.Add(listStyle);
                _listStyleByNumId[numId] = name;
                return name;
            }

            XElement BuildLevelStyle(int lvl, string fmt, string lvlText)
            {
                XElement props = new XElement(Style + "list-level-properties",
                    new XAttribute(Text + "list-level-position-and-space-mode", "label-alignment"),
                    new XElement(Style + "list-level-label-alignment",
                        new XAttribute(Text + "label-followed-by", "listtab"),
                        new XAttribute(Fo + "text-indent", "-0.635cm"),
                        new XAttribute(Fo + "margin-left", Cm((lvl + 1) * 1.27))));

                if (fmt == "bullet")
                {
                    string bullet = FirstNonSpaceChar(lvlText);
                    if (bullet == null) bullet = "•";
                    return new XElement(Text + "list-level-style-bullet",
                        new XAttribute(Text + "level", lvl + 1),
                        new XAttribute(Text + "bullet-char", bullet),
                        props);
                }
                string numFormat;
                if (fmt == "none")
                {
                    numFormat = "";                 // 空 num-format = 該層不顯示編號（Word numFmt="none"）
                }
                else
                {
                    bool recognized;
                    numFormat = MapNumFormat(fmt, out recognized);
                    if (!recognized && _warned.Add(fmt))
                        _warnings.Add("清單使用了尚未支援的編號格式「" + fmt +
                                      "」，已改以阿拉伯數字（1, 2, 3…）呈現，請產生後檢查。");
                }
                string prefix = NumPrefix(lvlText);
                string suffix = NumSuffix(lvlText);
                // ODF 的編號前後綴屬性在 style: 命名空間（不是 text:），否則無法通過 ODF 檢測。
                XElement el = new XElement(Text + "list-level-style-number",
                    new XAttribute(Text + "level", lvl + 1),
                    new XAttribute(Style + "num-format", numFormat));
                if (prefix != "") el.SetAttributeValue(Style + "num-prefix", prefix);
                if (suffix != "") el.SetAttributeValue(Style + "num-suffix", suffix);
                el.Add(props);
                return el;
            }

            // ================= 樣式檔（styles.xml，含頁面版面） =================
            public XDocument BuildStyles(XElement body)
            {
                XElement pageLayout = BuildPageLayout(body);

                XElement defaultPara = new XElement(Style + "default-style",
                    new XAttribute(Style + "family", "paragraph"),
                    new XElement(Style + "text-properties",
                        new XAttribute(Style + "font-name", DefaultWestern),
                        new XAttribute(Style + "font-name-asian", DefaultAsian),
                        new XAttribute(Style + "font-name-complex", DefaultWestern),
                        new XAttribute(Fo + "font-size", "12pt"),
                        new XAttribute(Style + "font-size-asian", "12pt"),
                        new XAttribute(Style + "font-size-complex", "12pt")));

                XElement standard = new XElement(Style + "style",
                    new XAttribute(Style + "name", "Standard"),
                    new XAttribute(Style + "family", "paragraph"),
                    new XAttribute(Style + "class", "text"));

                XElement root = new XElement(Office + "document-styles",
                    new XAttribute(XNamespace.Xmlns + "office", Office.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "style", Style.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "text", Text.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "fo", Fo.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "svg", Svg.NamespaceName),
                    new XAttribute(Office + "version", "1.2"),
                    BuildFontDecls(),
                    new XElement(Office + "styles", defaultPara, standard),
                    new XElement(Office + "automatic-styles", pageLayout),
                    new XElement(Office + "master-styles",
                        new XElement(Style + "master-page",
                            new XAttribute(Style + "name", "Standard"),
                            new XAttribute(Style + "page-layout-name", "PL1"))));
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            XElement BuildPageLayout(XElement body)
            {
                // A4 直式、2cm 邊界為預設；有 sectPr 就套用實際頁面設定
                double wCm = 21.0, hCm = 29.7, top = 2, bottom = 2, left = 2, right = 2;
                XElement sect = body.Descendants(W + "sectPr").LastOrDefault();
                if (sect != null)
                {
                    XElement pgSz = sect.Element(W + "pgSz");
                    if (pgSz != null)
                    {
                        double w = TwipsNum((string)pgSz.Attribute(W + "w"));
                        double h = TwipsNum((string)pgSz.Attribute(W + "h"));
                        if (w > 0) wCm = w; if (h > 0) hCm = h;
                    }
                    XElement pgMar = sect.Element(W + "pgMar");
                    if (pgMar != null)
                    {
                        top = TwipsNumOr(pgMar, "top", top);
                        bottom = TwipsNumOr(pgMar, "bottom", bottom);
                        left = TwipsNumOr(pgMar, "left", left);
                        right = TwipsNumOr(pgMar, "right", right);
                    }
                }
                XElement props = new XElement(Style + "page-layout-properties",
                    new XAttribute(Fo + "page-width", Cm(wCm)),
                    new XAttribute(Fo + "page-height", Cm(hCm)),
                    new XAttribute(Style + "print-orientation", wCm > hCm ? "landscape" : "portrait"),
                    new XAttribute(Fo + "margin-top", Cm(top)),
                    new XAttribute(Fo + "margin-bottom", Cm(bottom)),
                    new XAttribute(Fo + "margin-left", Cm(left)),
                    new XAttribute(Fo + "margin-right", Cm(right)));
                return new XElement(Style + "page-layout", new XAttribute(Style + "name", "PL1"), props);
            }

            XElement BuildFontDecls()
            {
                XElement decls = new XElement(Office + "font-face-decls");
                HashSet<string> all = new HashSet<string>(_fonts);
                all.Add(DefaultWestern);
                all.Add(DefaultAsian);
                foreach (string f in all.OrderBy(x => x, StringComparer.Ordinal))
                    decls.Add(new XElement(Style + "font-face",
                        new XAttribute(Style + "name", f),
                        new XAttribute(Svg + "font-family", QuoteFontIfNeeded(f))));
                return decls;
            }

            // ---- 清單判定小工具 ----
            string ListNumId(XElement p)
            {
                XElement numPr = NumPr(p);
                XElement numId = numPr != null ? numPr.Element(W + "numId") : null;
                if (numId != null)
                {
                    string v = (string)numId.Attribute(W + "val");
                    if (string.IsNullOrEmpty(v) || v == "0") return null;   // 直接 numId=0 = 取消編號
                    return v;
                }
                // 段落本身沒有直接 numPr → 編號可能來自段落樣式（w:pStyle，例如「清單編號」）
                string sv = _num.NumIdForStyle(PStyleId(p));
                if (!string.IsNullOrEmpty(sv) && sv != "0") return sv;
                return null;
            }
            int ListLevel(XElement p)
            {
                XElement numPr = NumPr(p);
                XElement ilvl = numPr != null ? numPr.Element(W + "ilvl") : null;
                int v;
                if (ilvl != null && int.TryParse((string)ilvl.Attribute(W + "val"), out v)) return v;
                return _num.IlvlForStyle(PStyleId(p));
            }
            static XElement NumPr(XElement p)
            {
                XElement pPr = p.Element(W + "pPr");
                return pPr != null ? pPr.Element(W + "numPr") : null;
            }
            static string PStyleId(XElement p)
            {
                XElement pPr = p.Element(W + "pPr");
                XElement ps = pPr != null ? pPr.Element(W + "pStyle") : null;
                return ps != null ? (string)ps.Attribute(W + "val") : null;
            }
        }

        // =================================================================
        //  numbering.xml 定義
        // =================================================================
        class NumberingDefs
        {
            public class Level { public string NumFmt = "bullet"; public string LvlText = ""; }
            public class Abstract
            {
                public readonly Dictionary<int, Level> Levels = new Dictionary<int, Level>();
                public Level Level(int lvl) { Level l; return Levels.TryGetValue(lvl, out l) ? l : null; }
            }

            readonly Dictionary<string, string> _numToAbstract = new Dictionary<string, string>();
            readonly Dictionary<string, Abstract> _abstracts = new Dictionary<string, Abstract>();
            readonly Dictionary<string, string> _styleNumId = new Dictionary<string, string>();  // pStyle → numId
            readonly Dictionary<string, int> _styleIlvl = new Dictionary<string, int>();          // pStyle → ilvl

            public static NumberingDefs Load(Dictionary<string, byte[]> parts)
            {
                NumberingDefs d = new NumberingDefs();
                LoadStyleNumbering(d, parts);   // 段落樣式帶來的編號（word/styles.xml）

                byte[] nb;
                if (!parts.TryGetValue("word/numbering.xml", out nb)) return d;
                XDocument xd;
                try { xd = ParseXml(nb); } catch { return d; }

                foreach (XElement a in xd.Descendants(W + "abstractNum"))
                {
                    string id = (string)a.Attribute(W + "abstractNumId");
                    if (id == null) continue;
                    Abstract abs = new Abstract();
                    foreach (XElement lvl in a.Elements(W + "lvl"))
                    {
                        int li;
                        if (!int.TryParse((string)lvl.Attribute(W + "ilvl"), out li)) continue;
                        Level lv = new Level();
                        XElement fmt = lvl.Element(W + "numFmt");
                        if (fmt != null) lv.NumFmt = (string)fmt.Attribute(W + "val") ?? "bullet";
                        XElement lt = lvl.Element(W + "lvlText");
                        if (lt != null) lv.LvlText = (string)lt.Attribute(W + "val") ?? "";
                        abs.Levels[li] = lv;
                    }
                    d._abstracts[id] = abs;
                }
                foreach (XElement n in xd.Descendants(W + "num"))
                {
                    string id = (string)n.Attribute(W + "numId");
                    XElement ab = n.Element(W + "abstractNumId");
                    if (id != null && ab != null) d._numToAbstract[id] = (string)ab.Attribute(W + "val");
                }
                return d;
            }

            public Abstract ForNumId(string numId)
            {
                string absId;
                if (numId == null || !_numToAbstract.TryGetValue(numId, out absId)) return null;
                Abstract abs;
                return absId != null && _abstracts.TryGetValue(absId, out abs) ? abs : null;
            }

            public string NumIdForStyle(string styleId)
            {
                string v;
                return styleId != null && _styleNumId.TryGetValue(styleId, out v) ? v : null;
            }
            public int IlvlForStyle(string styleId)
            {
                int v;
                return styleId != null && _styleIlvl.TryGetValue(styleId, out v) ? v : 0;
            }

            // 讀 word/styles.xml：把「段落樣式定義裡帶 numPr」的樣式記下來，讓只用 w:pStyle
            // 套用內建「清單編號／項目符號」樣式的段落也能被辨識為清單。
            static void LoadStyleNumbering(NumberingDefs d, Dictionary<string, byte[]> parts)
            {
                byte[] sb;
                if (!parts.TryGetValue("word/styles.xml", out sb)) return;
                XDocument sd;
                try { sd = ParseXml(sb); } catch { return; }
                foreach (XElement st in sd.Descendants(W + "style"))
                {
                    string sid = (string)st.Attribute(W + "styleId");
                    if (sid == null) continue;
                    XElement pPr = st.Element(W + "pPr");
                    XElement numPr = pPr != null ? pPr.Element(W + "numPr") : null;
                    if (numPr == null) continue;
                    XElement numId = numPr.Element(W + "numId");
                    string nv = numId != null ? (string)numId.Attribute(W + "val") : null;
                    if (string.IsNullOrEmpty(nv) || nv == "0") continue;
                    d._styleNumId[sid] = nv;
                    XElement ilvl = numPr.Element(W + "ilvl");
                    int lv;
                    if (ilvl != null && int.TryParse((string)ilvl.Attribute(W + "val"), out lv))
                        d._styleIlvl[sid] = lv;
                }
            }
        }

        // =================================================================
        //  固定 entry：meta / settings / manifest
        // =================================================================
        static XDocument BuildMeta()
        {
            XElement root = new XElement(Office + "document-meta",
                new XAttribute(XNamespace.Xmlns + "office", Office.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "meta", Meta.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "dc", Dc.NamespaceName),
                new XAttribute(Office + "version", "1.2"),
                new XElement(Office + "meta",
                    new XElement(Meta + "generator", "TenderDocGen DocxToOdt")));
            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        }

        static XDocument BuildSettings()
        {
            XElement root = new XElement(Office + "document-settings",
                new XAttribute(XNamespace.Xmlns + "office", Office.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "config", Config.NamespaceName),
                new XAttribute(Office + "version", "1.2"),
                new XElement(Office + "settings"));
            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        }

        static XDocument BuildManifest()
        {
            Func<string, string, XElement> fe = delegate(string path, string media)
            {
                return new XElement(Manifest + "file-entry",
                    new XAttribute(Manifest + "full-path", path),
                    new XAttribute(Manifest + "media-type", media));
            };
            XElement root = new XElement(Manifest + "manifest",
                new XAttribute(XNamespace.Xmlns + "manifest", Manifest.NamespaceName),
                new XAttribute(Manifest + "version", "1.2"),
                new XElement(Manifest + "file-entry",
                    new XAttribute(Manifest + "full-path", "/"),
                    new XAttribute(Manifest + "version", "1.2"),
                    new XAttribute(Manifest + "media-type", "application/vnd.oasis.opendocument.text")),
                fe("content.xml", "text/xml"),
                fe("styles.xml", "text/xml"),
                fe("meta.xml", "text/xml"),
                fe("settings.xml", "text/xml"));
            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        }

        // =================================================================
        //  小工具
        // =================================================================
        static KeyValuePair<string, byte[]> Entry(string name, byte[] data)
        {
            return new KeyValuePair<string, byte[]>(name, data);
        }

        static byte[] Serialize(XDocument doc)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                XmlWriterSettings s = new XmlWriterSettings();
                s.Encoding = new UTF8Encoding(false);
                s.Indent = false;
                using (XmlWriter xw = XmlWriter.Create(ms, s)) { doc.Save(xw); }
                return ms.ToArray();
            }
        }

        static Dictionary<string, byte[]> ReadParts(byte[] docx)
        {
            Dictionary<string, byte[]> d = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            using (MemoryStream ms = new MemoryStream(docx))
            using (ZipArchive zip = new ZipArchive(ms, ZipArchiveMode.Read))
                foreach (ZipArchiveEntry e in zip.Entries)
                {
                    using (MemoryStream mo = new MemoryStream())
                    using (Stream s = e.Open()) { s.CopyTo(mo); d[e.FullName] = mo.ToArray(); }
                }
            return d;
        }

        static XDocument ParseXml(byte[] bytes)
        {
            return XDocument.Parse(Encoding.UTF8.GetString(OdtWriter.StripBom(bytes)), LoadOptions.PreserveWhitespace);
        }

        static bool BoolProp(XElement rPr, string local)
        {
            XElement e = rPr.Element(W + local);
            if (e == null) return false;
            string v = (string)e.Attribute(W + "val");
            if (v == null) return true;                 // 無 val = 開啟
            v = v.Trim().ToLowerInvariant();
            return v != "false" && v != "0" && v != "off";
        }

        static bool IsHex6(string s)
        {
            if (s == null || s.Length != 6) return false;
            foreach (char c in s)
                if (!Uri.IsHexDigit(c)) return false;
            return true;
        }

        static string HalfPointToPt(string val)
        {
            int hp;
            if (val == null || !int.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out hp) || hp <= 0)
                return null;
            return (hp / 2.0).ToString("0.###", CultureInfo.InvariantCulture) + "pt";
        }

        static string MapAlign(string jc)
        {
            if (jc == null) return null;
            switch (jc)
            {
                case "left":
                case "start": return "start";
                case "right":
                case "end": return "end";
                case "center": return "center";
                case "both":
                case "distribute":
                case "justify": return "justify";
                default: return null;
            }
        }

        static string MapNumFormat(string fmt, out bool recognized)
        {
            recognized = true;
            switch (fmt)
            {
                case "decimal":
                case "decimalZero": return "1";
                case "lowerLetter": return "a";
                case "upperLetter": return "A";
                case "lowerRoman": return "i";
                case "upperRoman": return "I";
                default: recognized = false; return "1";   // CJK 等尚未支援的格式退回阿拉伯數字
            }
        }

        // lvlText 例如 "第%1條" → 前綴 "第"。取第一個 %n 之前的字元當前綴。
        static string NumPrefix(string lvlText)
        {
            if (string.IsNullOrEmpty(lvlText)) return "";
            int idx = lvlText.IndexOf('%');
            if (idx <= 0) return "";
            return lvlText.Substring(0, idx);
        }

        // lvlText 例如 "%1." → 後綴 "."；"%1)" → ")"。取 %n 之後的字元當後綴。
        static string NumSuffix(string lvlText)
        {
            if (string.IsNullOrEmpty(lvlText)) return "";
            int idx = lvlText.LastIndexOf('%');
            if (idx < 0 || idx + 1 >= lvlText.Length) return "";
            // 跳過 % 後的數字
            int j = idx + 1;
            while (j < lvlText.Length && char.IsDigit(lvlText[j])) j++;
            return lvlText.Substring(j);
        }

        static string FirstNonSpaceChar(string s)
        {
            if (string.IsNullOrEmpty(s)) return null;
            foreach (char c in s)
                if (c != ' ' && c != '\t') return c.ToString();
            return null;
        }

        static string QuoteFontIfNeeded(string f)
        {
            if (f != null && (f.Contains(" ") || f.Contains(","))) return "'" + f + "'";
            return f;
        }

        // ---- 單位換算 ----
        static string Cm(double cm) { return cm.ToString("0.###", CultureInfo.InvariantCulture) + "cm"; }

        static string TwipsToCm(string twips)
        {
            int tw;
            if (!int.TryParse(twips, NumberStyles.Integer, CultureInfo.InvariantCulture, out tw)) return "0cm";
            return Cm(tw / 1440.0 * 2.54);
        }

        // pgSz/pgMar 用的：twips → cm（數值）
        static double TwipsNum(string twips)
        {
            int tw;
            if (twips == null || !int.TryParse(twips, NumberStyles.Integer, CultureInfo.InvariantCulture, out tw)) return 0;
            return tw / 1440.0 * 2.54;
        }
        static double TwipsNumOr(XElement el, string local, double fallback)
        {
            double v = TwipsNum((string)el.Attribute(W + local));
            return v > 0 ? v : fallback;
        }

        // w:ind 的 left/right 可能用新式 start/end 屬性名
        static string TwipsAttr(XElement ind, string legacy, string modern)
        {
            string v = (string)ind.Attribute(W + legacy);
            if (string.IsNullOrEmpty(v)) v = (string)ind.Attribute(W + modern);
            if (string.IsNullOrEmpty(v)) return null;
            return TwipsToCm(v);
        }
    }
}
