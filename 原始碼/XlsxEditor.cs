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
    /// <summary>要對 標案資料.xlsx 追加的欄／列。</summary>
    class XlsxAdditions
    {
        public List<string> TenderColumns = new List<string>();        // 附加到「標案清單」表右側（依序）
        public HashSet<string> DropdownColumns = new HashSet<string>(); // 其中要套「是/否」下拉的欄名
        public List<string> CompanyParams = new List<string>();        // 附加到「公司資料」的參數名

        public bool IsEmpty { get { return TenderColumns.Count == 0 && CompanyParams.Count == 0; } }
    }

    /// <summary>
    /// 以純 zip + XML 對既有 標案資料.xlsx 追加欄／列（不依賴 Office）。
    /// 「標案清單」是 ListObject 表格：同步更新 table 的 ref/autoFilter/tableColumns 與
    /// worksheet 的表頭儲存格、dimension、（產生欄）dataValidation。
    /// 寫入採「暫存檔→驗證→取代」並先備份，任何失敗都不動到原檔。
    /// </summary>
    static class XlsxEditor
    {
        static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string SheetTenders = "標案清單";
        const string SheetCompany = "公司資料";

        public static void Apply(string xlsxPath, XlsxAdditions add)
        {
            if (add == null || add.IsEmpty) return;
            if (!File.Exists(xlsxPath))
                throw new FileNotFoundException("找不到 標案資料.xlsx：" + xlsxPath);

            // 產生前先記錄既有標案列數，供事後比對（確保沒破壞資料）
            int beforeTenderRows = CountTenderDataRows(xlsxPath);

            List<KeyValuePair<string, byte[]>> entries = OdtWriter.ReadAllEntries(xlsxPath);

            // 解析工作表名稱 → worksheet entry 路徑
            Dictionary<string, string> sheetPaths = ResolveSheetPaths(entries);
            // Excel 存檔時表頭多為 sharedString(t="s")，需先載入才能正確比對既有欄名（避免重複加欄）
            List<string> shared = LoadSharedStrings(entries);

            if (add.TenderColumns.Count > 0)
            {
                string wsPath;
                if (!sheetPaths.TryGetValue(SheetTenders, out wsPath))
                    throw new InvalidDataException("Excel 內找不到「" + SheetTenders + "」工作表。");
                string tablePath = ResolveTablePath(entries, wsPath);
                AppendTenderColumns(entries, wsPath, tablePath, add, shared);
            }

            if (add.CompanyParams.Count > 0)
            {
                string wsPath;
                if (!sheetPaths.TryGetValue(SheetCompany, out wsPath))
                    throw new InvalidDataException("Excel 內找不到「" + SheetCompany + "」工作表。");
                AppendCompanyRows(entries, wsPath, add.CompanyParams, shared);
            }

            byte[] bytes = OdtWriter.Repack(entries);

            // 備份 → 暫存檔 → 驗證 → 取代
            string backup = xlsxPath + ".備份-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".xlsx";
            string tmp = xlsxPath + ".tmp";
            File.Copy(xlsxPath, backup, true);
            try
            {
                File.WriteAllBytes(tmp, bytes);
                ValidateResult(tmp, add, beforeTenderRows);   // 失敗會丟例外

                try
                {
                    // 原子取代：tmp→xlsxPath（失敗時原檔應保留）
                    File.Replace(tmp, xlsxPath, null);
                }
                catch (Exception ex)
                {
                    // 保險：萬一原檔被移走而取代失敗，用備份還原
                    try { if (!File.Exists(xlsxPath) && File.Exists(backup)) File.Copy(backup, xlsxPath, true); }
                    catch { }
                    if (ex is IOException || ex is UnauthorizedAccessException)
                        throw new IOException("無法更新 標案資料.xlsx —— 它可能正被 Excel 開啟，請關閉後再試一次。");
                    throw;
                }
                // 成功：移除備份，避免堆積
                try { File.Delete(backup); } catch { }
            }
            catch
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
                // 原檔仍在就順手清掉備份；原檔不在則保留備份供人工還原
                try { if (File.Exists(xlsxPath)) File.Delete(backup); } catch { }
                throw;
            }
        }

        // ---------- 標案清單：加欄 ----------

        static void AppendTenderColumns(List<KeyValuePair<string, byte[]>> entries,
            string wsPath, string tablePath, XlsxAdditions add, List<string> shared)
        {
            XDocument ws = LoadPart(entries, wsPath);
            XDocument tbl = tablePath != null ? LoadPart(entries, tablePath) : null;

            XElement sheetData = ws.Descendants(Main + "sheetData").First();
            XElement headerRow = sheetData.Elements(Main + "row").FirstOrDefault();
            if (headerRow == null) throw new InvalidDataException("「標案清單」缺少表頭列。");

            // 既有表頭名稱（含 sharedString），用來避免重複加欄；
            // 也把 table 既有的 tableColumn 名一併納入（雙重保險，ListObject 欄名不可重複）
            HashSet<string> existingHeaders = new HashSet<string>(
                headerRow.Elements(Main + "c").Select(c => Util.Nfkc(CellText(c, shared))).Where(s => s != ""));
            if (tbl != null)
                foreach (XElement tc in tbl.Descendants(Main + "tableColumn"))
                {
                    string nm = Util.Nfkc((string)tc.Attribute("name"));
                    if (nm != "") existingHeaders.Add(nm);
                }

            // 目前寬度與最後資料列
            int lastCol = 0, lastRow = 1;
            foreach (XElement r in sheetData.Elements(Main + "row"))
            {
                int rn; int.TryParse((string)r.Attribute("r"), out rn);
                if (rn > lastRow) lastRow = rn;
                foreach (XElement c in r.Elements(Main + "c"))
                {
                    int ci = ColIndex((string)c.Attribute("r"));
                    if (ci > lastCol) lastCol = ci;
                }
            }

            // 表頭儲存格樣式沿用第一格
            string headerStyle = (string)headerRow.Elements(Main + "c").First().Attribute("s");

            List<XElement> validationRanges = new List<XElement>();
            int addedTo = lastCol;
            List<string> reallyAdded = new List<string>();
            HashSet<string> addedThisRun = new HashSet<string>();
            foreach (string name in add.TenderColumns)
            {
                string nn = Util.Nfkc(name);
                if (existingHeaders.Contains(nn) || addedThisRun.Contains(nn)) continue; // 已存在/本次已加 → 略過
                addedThisRun.Add(nn);
                addedTo++;
                string colL = ColLetter(addedTo);
                XElement hc = new XElement(Main + "c",
                    new XAttribute("r", colL + "1"),
                    new XElement(Main + "is", new XElement(Main + "t", name)));
                hc.SetAttributeValue("t", "inlineStr");
                if (!string.IsNullOrEmpty(headerStyle)) hc.SetAttributeValue("s", headerStyle);
                headerRow.Add(hc);
                reallyAdded.Add(name);

                if (add.DropdownColumns.Contains(name))
                {
                    // lastRow<2（僅表頭）時避免產生起訖顛倒的 sqref
                    string sq = lastRow >= 2 ? (colL + "2:" + colL + lastRow) : (colL + "2");
                    validationRanges.Add(new XElement("range", sq));
                }
            }
            if (reallyAdded.Count == 0) { SavePart(entries, wsPath, ws); return; }

            int newLastCol = addedTo;

            // 更新 dimension
            XElement dim = ws.Descendants(Main + "dimension").FirstOrDefault();
            if (dim != null)
            {
                dim.SetAttributeValue("ref", "A1:" + ColLetter(newLastCol) + lastRow);
            }
            // 更新表頭列 spans（best-effort）
            if (headerRow.Attribute("spans") != null)
                headerRow.SetAttributeValue("spans", "1:" + newLastCol);

            // 產生欄的「是/否」下拉
            if (validationRanges.Count > 0)
            {
                XElement dvs = ws.Descendants(Main + "dataValidations").FirstOrDefault();
                if (dvs == null)
                {
                    dvs = new XElement(Main + "dataValidations", new XAttribute("count", 0));
                    // 需置於 sheetData 之後、pageMargins 之前
                    XElement pm = ws.Descendants(Main + "pageMargins").FirstOrDefault();
                    if (pm != null) pm.AddBeforeSelf(dvs); else sheetData.AddAfterSelf(dvs);
                }
                foreach (XElement rng in validationRanges)
                {
                    XElement dv = new XElement(Main + "dataValidation",
                        new XAttribute("type", "list"),
                        new XAttribute("allowBlank", "1"),
                        new XAttribute("showInputMessage", "1"),
                        new XAttribute("showErrorMessage", "1"),
                        new XAttribute("sqref", rng.Value),
                        new XElement(Main + "formula1", "\"是,否\""));
                    dvs.Add(dv);
                }
                int cnt = dvs.Elements(Main + "dataValidation").Count();
                dvs.SetAttributeValue("count", cnt);
            }

            SavePart(entries, wsPath, ws);

            // 同步 table：ref / autoFilter / tableColumns
            if (tbl != null)
            {
                XElement tableEl = tbl.Root;
                string oldRef = (string)tableEl.Attribute("ref");           // 例 A1:L4
                string lastRowOfTable = Regex.Match(oldRef, @":\D+(\d+)$").Groups[1].Value;
                string newRef = "A1:" + ColLetter(newLastCol) + lastRowOfTable;
                tableEl.SetAttributeValue("ref", newRef);
                XElement af = tableEl.Element(Main + "autoFilter");
                if (af != null) af.SetAttributeValue("ref", newRef);

                XElement cols = tableEl.Element(Main + "tableColumns");
                int maxId = cols.Elements(Main + "tableColumn")
                    .Select(c => { int v; int.TryParse((string)c.Attribute("id"), out v); return v; })
                    .DefaultIfEmpty(0).Max();
                foreach (string name in reallyAdded)
                {
                    maxId++;
                    cols.Add(new XElement(Main + "tableColumn",
                        new XAttribute("id", maxId),
                        new XAttribute("name", name)));
                }
                cols.SetAttributeValue("count", cols.Elements(Main + "tableColumn").Count());
                SavePart(entries, tablePath, tbl);
            }
        }

        // ---------- 公司資料：加列 ----------

        static void AppendCompanyRows(List<KeyValuePair<string, byte[]>> entries,
            string wsPath, List<string> paramNames, List<string> shared)
        {
            XDocument ws = LoadPart(entries, wsPath);
            XElement sheetData = ws.Descendants(Main + "sheetData").First();

            // 既有參數名（A 欄，含 sharedString），避免重複
            HashSet<string> existing = new HashSet<string>();
            int lastRow = 0;
            foreach (XElement r in sheetData.Elements(Main + "row"))
            {
                int rn; int.TryParse((string)r.Attribute("r"), out rn);
                if (rn > lastRow) lastRow = rn;
                XElement a = r.Elements(Main + "c").FirstOrDefault(c => ColIndex((string)c.Attribute("r")) == 1);
                if (a != null) { string t = Util.Nfkc(CellText(a, shared)); if (t != "") existing.Add(t); }
            }

            List<string> added = new List<string>();
            foreach (string name in paramNames)
            {
                if (existing.Contains(Util.Nfkc(name))) continue;
                lastRow++;
                XElement row = new XElement(Main + "row", new XAttribute("r", lastRow));
                row.Add(new XElement(Main + "c",
                    new XAttribute("r", "A" + lastRow),
                    new XAttribute("t", "inlineStr"),
                    new XElement(Main + "is", new XElement(Main + "t", name))));
                row.Add(new XElement(Main + "c",
                    new XAttribute("r", "C" + lastRow),
                    new XAttribute("t", "inlineStr"),
                    new XElement(Main + "is", new XElement(Main + "t", "（新範本需要，請填入固定值）"))));
                sheetData.Add(row);
                added.Add(name);
            }
            if (added.Count == 0) { SavePart(entries, wsPath, ws); return; }

            XElement dim = ws.Descendants(Main + "dimension").FirstOrDefault();
            if (dim != null)
            {
                string oldRef = (string)dim.Attribute("ref");
                int lastColLetter = 3; // 公司資料固定 3 欄 A:C
                dim.SetAttributeValue("ref", "A1:" + ColLetter(lastColLetter) + lastRow);
            }
            SavePart(entries, wsPath, ws);
        }

        // ---------- 解析 ----------

        static Dictionary<string, string> ResolveSheetPaths(List<KeyValuePair<string, byte[]>> entries)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            XDocument wb = LoadPart(entries, "xl/workbook.xml");
            XDocument rels = LoadPartOrNull(entries, "xl/_rels/workbook.xml.rels");
            Dictionary<string, string> relTarget = new Dictionary<string, string>();
            if (rels != null)
                foreach (XElement rel in rels.Descendants().Where(e => e.Name.LocalName == "Relationship"))
                {
                    string id = (string)rel.Attribute("Id");
                    string target = (string)rel.Attribute("Target");
                    if (id != null && target != null) relTarget[id] = ResolveXlPath(target);
                }

            int i = 0;
            foreach (XElement sheet in wb.Descendants().Where(e => e.Name.LocalName == "sheet"))
            {
                i++;
                string name = (string)sheet.Attribute("name");
                string rid = sheet.Attributes().Where(a => a.Name.LocalName == "id")
                    .Select(a => a.Value).FirstOrDefault();
                string path = null;
                if (rid != null && relTarget.ContainsKey(rid)) path = relTarget[rid];
                if (path == null) path = "xl/worksheets/sheet" + i + ".xml";
                if (name != null) result[Util.Nfkc(name)] = path;
            }
            return result;
        }

        static string ResolveTablePath(List<KeyValuePair<string, byte[]>> entries, string wsPath)
        {
            XDocument ws = LoadPart(entries, wsPath);
            XElement tp = ws.Descendants(Main + "tablePart").FirstOrDefault();
            if (tp == null) return null;
            string rid = tp.Attributes().Where(a => a.Name.LocalName == "id").Select(a => a.Value).FirstOrDefault();
            if (rid == null) return null;

            string relsPath = "xl/worksheets/_rels/" + Path.GetFileName(wsPath) + ".rels";
            XDocument rels = LoadPartOrNull(entries, relsPath);
            if (rels == null) return null;
            foreach (XElement rel in rels.Descendants().Where(e => e.Name.LocalName == "Relationship"))
            {
                if ((string)rel.Attribute("Id") != rid) continue;
                string target = (string)rel.Attribute("Target");   // ../tables/table1.xml
                return NormalizeRelative(wsPath, target);
            }
            return null;
        }

        // ---------- 驗證 ----------

        static int CountTenderDataRows(string xlsxPath)
        {
            try
            {
                Dictionary<string, List<XlsxRow>> book = XlsxReader.Load(xlsxPath);
                List<XlsxRow> rows;
                foreach (KeyValuePair<string, List<XlsxRow>> kv in book)
                    if (Util.Nfkc(kv.Key) == SheetTenders)
                    {
                        rows = kv.Value;
                        return rows.Skip(1).Count(r => r.Cells.Any(c => (c ?? "").Trim() != ""));
                    }
            }
            catch { }
            return -1;
        }

        static void ValidateResult(string path, XlsxAdditions add, int beforeTenderRows)
        {
            Dictionary<string, List<XlsxRow>> book;
            try { book = XlsxReader.Load(path); }
            catch (Exception ex) { throw new InvalidDataException("更新後的 Excel 無法讀取：" + ex.Message); }

            List<XlsxRow> tenderRows = null, companyRows = null;
            foreach (KeyValuePair<string, List<XlsxRow>> kv in book)
            {
                if (Util.Nfkc(kv.Key) == SheetTenders) tenderRows = kv.Value;
                if (Util.Nfkc(kv.Key) == SheetCompany) companyRows = kv.Value;
            }

            if (add.TenderColumns.Count > 0)
            {
                if (tenderRows == null || tenderRows.Count == 0)
                    throw new InvalidDataException("更新後找不到「標案清單」表頭。");
                HashSet<string> headers = new HashSet<string>(
                    tenderRows[0].Cells.Select(c => Util.Nfkc(c)));
                foreach (string col in add.TenderColumns)
                    if (!headers.Contains(Util.Nfkc(col)))
                        throw new InvalidDataException("更新後缺少欄位「" + col + "」，已還原原檔。");

                int after = tenderRows.Skip(1).Count(r => r.Cells.Any(c => (c ?? "").Trim() != ""));
                if (beforeTenderRows >= 0 && after != beforeTenderRows)
                    throw new InvalidDataException("更新後標案資料列數改變（" + beforeTenderRows + "→" + after + "），已還原原檔。");
            }

            if (add.CompanyParams.Count > 0)
            {
                if (companyRows == null) throw new InvalidDataException("更新後找不到「公司資料」工作表。");
                HashSet<string> keys = new HashSet<string>(
                    companyRows.Select(r => Util.Nfkc(r.Cell(0))));
                foreach (string p in add.CompanyParams)
                    if (!keys.Contains(Util.Nfkc(p)))
                        throw new InvalidDataException("更新後缺少公司參數「" + p + "」，已還原原檔。");
            }
        }

        // ---------- zip / XML part 存取 ----------

        static XDocument LoadPart(List<KeyValuePair<string, byte[]>> entries, string name)
        {
            XDocument d = LoadPartOrNull(entries, name);
            if (d == null) throw new InvalidDataException("Excel 內缺少必要的部件：" + name);
            return d;
        }

        static XDocument LoadPartOrNull(List<KeyValuePair<string, byte[]>> entries, string name)
        {
            int idx = entries.FindIndex(e => string.Equals(e.Key, name, StringComparison.OrdinalIgnoreCase));
            if (idx < 0) return null;
            return XDocument.Parse(Encoding.UTF8.GetString(OdtWriter.StripBom(entries[idx].Value)),
                LoadOptions.PreserveWhitespace);
        }

        static void SavePart(List<KeyValuePair<string, byte[]>> entries, string name, XDocument doc)
        {
            int idx = entries.FindIndex(e => string.Equals(e.Key, name, StringComparison.OrdinalIgnoreCase));
            if (idx < 0) throw new InvalidDataException("找不到要寫回的部件：" + name);
            if (doc.Declaration == null) doc.Declaration = new XDeclaration("1.0", "UTF-8", "yes");
            byte[] bytes;
            using (MemoryStream ms = new MemoryStream())
            {
                XmlWriterSettings s = new XmlWriterSettings();
                s.Encoding = new UTF8Encoding(false);
                s.Indent = false;
                using (XmlWriter xw = XmlWriter.Create(ms, s)) { doc.Save(xw); }
                bytes = ms.ToArray();
            }
            entries[idx] = new KeyValuePair<string, byte[]>(entries[idx].Key, bytes);
        }

        // ---------- 小工具 ----------

        static string CellText(XElement c, List<string> shared)
        {
            string t = (string)c.Attribute("t");
            if (t == "inlineStr")
                return string.Concat(c.Descendants(Main + "t").Select(x => x.Value));
            XElement v = c.Element(Main + "v");
            if (t == "s")
            {
                int idx;
                if (v != null && int.TryParse(v.Value, out idx) && idx >= 0 && shared != null && idx < shared.Count)
                    return shared[idx];
                return "";
            }
            return v != null ? v.Value : "";
        }

        /// <summary>載入 xl/sharedStrings.xml（排除 rPh 注音），回傳依索引的字串清單。</summary>
        static List<string> LoadSharedStrings(List<KeyValuePair<string, byte[]>> entries)
        {
            List<string> list = new List<string>();
            XDocument sst = LoadPartOrNull(entries, "xl/sharedStrings.xml");
            if (sst == null) return list;
            foreach (XElement si in sst.Descendants(Main + "si"))
            {
                StringBuilder sb = new StringBuilder();
                foreach (XElement tEl in si.Descendants(Main + "t"))
                {
                    bool phonetic = false;
                    for (XElement p = tEl.Parent; p != null && p != si; p = p.Parent)
                        if (p.Name.LocalName == "rPh") { phonetic = true; break; }
                    if (!phonetic) sb.Append(tEl.Value);
                }
                list.Add(sb.ToString());
            }
            return list;
        }

        static int ColIndex(string cellRef)
        {
            if (cellRef == null) return 0;
            int col = 0;
            foreach (char ch in cellRef)
            {
                char u = char.ToUpperInvariant(ch);
                if (u >= 'A' && u <= 'Z') col = col * 26 + (u - 'A' + 1);
                else break;
            }
            return col;
        }

        static string ColLetter(int index1)
        {
            string s = "";
            int n = index1;
            while (n > 0) { int r = (n - 1) % 26; s = (char)('A' + r) + s; n = (n - 1) / 26; }
            return s;
        }

        static string ResolveXlPath(string target)
        {
            target = target.Replace('\\', '/');
            if (target.StartsWith("/")) return target.TrimStart('/');
            if (target.StartsWith("xl/")) return target;
            return "xl/" + target;
        }

        static string NormalizeRelative(string basePartPath, string relative)
        {
            // basePartPath 例：xl/worksheets/sheet1.xml；relative 例：../tables/table1.xml
            relative = relative.Replace('\\', '/');
            if (relative.StartsWith("/")) return relative.TrimStart('/');
            List<string> baseParts = basePartPath.Split('/').ToList();
            baseParts.RemoveAt(baseParts.Count - 1);   // 去掉檔名
            foreach (string seg in relative.Split('/'))
            {
                if (seg == "..") { if (baseParts.Count > 0) baseParts.RemoveAt(baseParts.Count - 1); }
                else if (seg == "." || seg == "") { }
                else baseParts.Add(seg);
            }
            return string.Join("/", baseParts);
        }
    }
}
