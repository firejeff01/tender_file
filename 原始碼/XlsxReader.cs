using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace TenderDocGen
{
    /// <summary>xlsx 工作表的一列。RowNumber 為 Excel 實際列號（1 起算）。</summary>
    class XlsxRow
    {
        public int RowNumber;
        public List<string> Cells = new List<string>();

        public string Cell(int index)
        {
            if (index < 0 || index >= Cells.Count) return "";
            return Cells[index] ?? "";
        }
    }

    /// <summary>
    /// 純 zip+XML 的 xlsx 讀取器，執行期不依賴任何 Office。
    /// 相容：sharedStrings（含 rich-text 多段）、inlineStr、公式快取字串、
    /// r 屬性缺漏（依序推進）、跳號儲存格、LibreOffice Calc 另存的檔案。
    /// </summary>
    static class XlsxReader
    {
        /// <summary>讀取整本活頁簿：工作表名稱 → 列集合。以 FileShare.ReadWrite 開啟（Excel 開著也能讀）。</summary>
        public static Dictionary<string, List<XlsxRow>> Load(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("找不到 Excel 檔案：" + path);

            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    return LoadFromZip(zip);
                }
            }
            catch (InvalidDataException)
            {
                throw new InvalidDataException(
                    "無法讀取「" + Path.GetFileName(path) + "」：檔案不是 .xlsx 格式。\n" +
                    "若曾用 LibreOffice 另存，請確認存檔類型選「Excel 2007-365 (.xlsx)」而不是 .ods。");
            }
        }

        static Dictionary<string, List<XlsxRow>> LoadFromZip(ZipArchive zip)
        {
            // 1. workbook.xml：工作表清單（名稱 + rId）
            XDocument workbook = LoadXml(zip, "xl/workbook.xml");
            if (workbook == null) throw new InvalidDataException("xlsx 缺少 xl/workbook.xml");

            // 2. workbook 關聯：rId → 目標路徑
            Dictionary<string, string> relTargets = new Dictionary<string, string>();
            XDocument rels = LoadXml(zip, "xl/_rels/workbook.xml.rels");
            if (rels != null)
            {
                foreach (XElement rel in Descendants(rels, "Relationship"))
                {
                    string id = AttrByLocalName(rel, "Id");
                    string target = AttrByLocalName(rel, "Target");
                    if (id != null && target != null) relTargets[id] = ResolveTarget(target);
                }
            }

            // 3. sharedStrings（可能不存在）
            List<string> shared = new List<string>();
            XDocument sst = LoadXml(zip, "xl/sharedStrings.xml");
            if (sst != null)
            {
                foreach (XElement si in Descendants(sst, "si"))
                    shared.Add(ConcatTexts(si));
            }

            // 4. 各工作表
            Dictionary<string, List<XlsxRow>> result = new Dictionary<string, List<XlsxRow>>();
            int sheetIndex = 0;
            foreach (XElement sheet in Descendants(workbook, "sheet"))
            {
                sheetIndex++;
                string name = AttrByLocalName(sheet, "name") ?? ("Sheet" + sheetIndex);
                string rid = null;
                foreach (XAttribute a in sheet.Attributes())
                    if (a.Name.LocalName == "id" && a.Name.Namespace != XNamespace.None) rid = a.Value;

                string target = null;
                if (rid != null && relTargets.ContainsKey(rid)) target = relTargets[rid];
                if (target == null) target = "xl/worksheets/sheet" + sheetIndex + ".xml"; // 後備

                XDocument ws = LoadXml(zip, target);
                if (ws == null) continue;
                result[name] = ParseSheet(ws, shared);
            }
            return result;
        }

        static List<XlsxRow> ParseSheet(XDocument ws, List<string> shared)
        {
            List<XlsxRow> rows = new List<XlsxRow>();
            int lastRowNum = 0;
            foreach (XElement rowEl in Descendants(ws, "row"))
            {
                int rowNum;
                string rAttr = AttrByLocalName(rowEl, "r");
                if (rAttr == null || !int.TryParse(rAttr, out rowNum)) rowNum = lastRowNum + 1;
                lastRowNum = rowNum;

                XlsxRow row = new XlsxRow();
                row.RowNumber = rowNum;
                int lastCol = -1;
                foreach (XElement c in rowEl.Elements().Where(e => e.Name.LocalName == "c"))
                {
                    int col;
                    string cellRef = AttrByLocalName(c, "r");
                    if (cellRef != null) col = ColumnIndexFromRef(cellRef);
                    else col = lastCol + 1;
                    lastCol = col;

                    string value = CellValue(c, shared);
                    while (row.Cells.Count <= col) row.Cells.Add("");
                    row.Cells[col] = value;
                }
                rows.Add(row);
            }
            return rows;
        }

        static string CellValue(XElement c, List<string> shared)
        {
            string t = AttrByLocalName(c, "t") ?? "";
            XElement v = c.Elements().FirstOrDefault(e => e.Name.LocalName == "v");
            XElement isEl = c.Elements().FirstOrDefault(e => e.Name.LocalName == "is");

            if (t == "inlineStr")
                return isEl != null ? ConcatTexts(isEl) : "";
            if (v == null) return "";

            string raw = v.Value;
            if (t == "s")
            {
                int idx;
                if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out idx) &&
                    idx >= 0 && idx < shared.Count) return shared[idx];
                return "";
            }
            if (t == "str") return raw;              // 公式快取字串
            if (t == "b") return raw == "1" ? "是" : "否";
            if (t == "e") return "";                 // 錯誤值（#N/A 等）視為空白
            return Util.ExpandScientific(raw);       // 數字（防科學記號）
        }

        /// <summary>
        /// 串接元素底下所有 t 節點文字（rich-text si 會拆成多個 r/t）。
        /// 排除 rPh（注音／振假名導引）內的 t，否則亞洲語系儲存格值會被注音污染。
        /// </summary>
        static string ConcatTexts(XElement el)
        {
            StringBuilder sb = new StringBuilder();
            foreach (XElement tEl in el.Descendants().Where(e => e.Name.LocalName == "t"))
            {
                bool insidePhonetic = false;
                for (XElement p = tEl.Parent; p != null; p = p.Parent)
                    if (p.Name.LocalName == "rPh") { insidePhonetic = true; break; }
                if (!insidePhonetic) sb.Append(tEl.Value);
            }
            return sb.ToString();
        }

        /// <summary>「BC12」→ 54（0 起算欄索引）。</summary>
        static int ColumnIndexFromRef(string cellRef)
        {
            int col = 0;
            foreach (char ch in cellRef)
            {
                if (ch >= 'A' && ch <= 'Z') col = col * 26 + (ch - 'A' + 1);
                else if (ch >= 'a' && ch <= 'z') col = col * 26 + (ch - 'a' + 1);
                else break;
            }
            return col - 1;
        }

        static string ResolveTarget(string target)
        {
            target = target.Replace('\\', '/');
            if (target.StartsWith("/")) return target.TrimStart('/');
            if (target.StartsWith("xl/")) return target;
            return "xl/" + target;
        }

        static XDocument LoadXml(ZipArchive zip, string entryName)
        {
            ZipArchiveEntry entry = zip.GetEntry(entryName);
            if (entry == null)
            {
                // 大小寫不敏感後備（某些工具產出的 zip 大小寫不一）
                entry = zip.Entries.FirstOrDefault(
                    e => string.Equals(e.FullName, entryName, StringComparison.OrdinalIgnoreCase));
            }
            if (entry == null) return null;
            using (Stream s = entry.Open())
                return XDocument.Load(s);
        }

        /// <summary>依 LocalName 找所有後代（不綁定 namespace，相容 Strict OOXML / LO 產出）。</summary>
        static IEnumerable<XElement> Descendants(XDocument doc, string localName)
        {
            return doc.Descendants().Where(e => e.Name.LocalName == localName);
        }

        static string AttrByLocalName(XElement el, string localName)
        {
            foreach (XAttribute a in el.Attributes())
                if (a.Name.LocalName == localName) return a.Value;
            return null;
        }
    }
}
