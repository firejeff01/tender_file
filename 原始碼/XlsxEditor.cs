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
        public List<string> TenderColumns = new List<string>();        // 依序插入「標案清單」（產生欄群組之後）
        public HashSet<string> DropdownColumns = new HashSet<string>(); // 其中要套「是/否」下拉的欄名
        public List<string> CompanyParams = new List<string>();        // 附加到「公司資料」的參數名
        public bool IsEmpty { get { return TenderColumns.Count == 0 && CompanyParams.Count == 0; } }
    }

    /// <summary>要從 標案資料.xlsx 移除的欄／列。</summary>
    class XlsxRemovals
    {
        public List<string> TenderColumns = new List<string>();        // 從「標案清單」移除的欄名
        public List<string> CompanyParams = new List<string>();        // 從「公司資料」移除的參數名
        public bool IsEmpty { get { return TenderColumns.Count == 0 && CompanyParams.Count == 0; } }
    }

    /// <summary>
    /// 以純 zip + XML 對既有 標案資料.xlsx 插入／移除欄列（不依賴 Office）。
    /// 「標案清單」是 ListObject 表格：同步維護 table 的 ref/autoFilter/tableColumns 與
    /// worksheet 的儲存格參照、dimension、dataValidation。
    /// 寫入採「備份→暫存→XlsxReader 重讀驗證→File.Replace 原子取代」，任何失敗都不動原檔。
    /// </summary>
    static class XlsxEditor
    {
        static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string SheetTenders = "標案清單";
        const string SheetCompany = "公司資料";

        // ================= 對外：插入 =================
        public static void Apply(string xlsxPath, XlsxAdditions add)
        {
            if (add == null || add.IsEmpty) return;
            if (!File.Exists(xlsxPath)) throw new FileNotFoundException("找不到 標案資料.xlsx：" + xlsxPath);

            int beforeRows = CountTenderDataRows(xlsxPath);
            List<KeyValuePair<string, byte[]>> entries = OdtWriter.ReadAllEntries(xlsxPath);
            Dictionary<string, string> sheetPaths = ResolveSheetPaths(entries);
            List<string> shared = LoadSharedStrings(entries);

            if (add.TenderColumns.Count > 0)
            {
                string wsPath = RequireSheet(sheetPaths, SheetTenders);
                InsertTenderColumns(entries, wsPath, ResolveTablePath(entries, wsPath), add, shared);
            }
            if (add.CompanyParams.Count > 0)
                AppendCompanyRows(entries, RequireSheet(sheetPaths, SheetCompany), add.CompanyParams, shared);

            WriteAndValidate(xlsxPath, entries, delegate(string path)
            {
                Dictionary<string, List<XlsxRow>> book = ReopenAndCheck(path);
                if (add.TenderColumns.Count > 0)
                {
                    HashSet<string> hdr = TenderHeaders(book);
                    foreach (string col in add.TenderColumns)
                        if (!hdr.Contains(Util.Nfkc(col)))
                            throw new InvalidDataException("更新後缺少欄位「" + col + "」，已還原原檔。");
                    CheckRowCount(book, beforeRows);
                }
                if (add.CompanyParams.Count > 0)
                {
                    HashSet<string> keys = CompanyKeys(book);
                    foreach (string p in add.CompanyParams)
                        if (!keys.Contains(Util.Nfkc(p)))
                            throw new InvalidDataException("更新後缺少公司參數「" + p + "」，已還原原檔。");
                }
            });
        }

        // ================= 對外：移除 =================
        public static void Remove(string xlsxPath, XlsxRemovals rem)
        {
            if (rem == null || rem.IsEmpty) return;
            if (!File.Exists(xlsxPath)) throw new FileNotFoundException("找不到 標案資料.xlsx：" + xlsxPath);

            int beforeRows = CountTenderDataRows(xlsxPath);
            List<KeyValuePair<string, byte[]>> entries = OdtWriter.ReadAllEntries(xlsxPath);
            Dictionary<string, string> sheetPaths = ResolveSheetPaths(entries);
            List<string> shared = LoadSharedStrings(entries);

            if (rem.TenderColumns.Count > 0)
            {
                string wsPath = RequireSheet(sheetPaths, SheetTenders);
                RemoveTenderColumns(entries, wsPath, ResolveTablePath(entries, wsPath), rem.TenderColumns, shared);
            }
            if (rem.CompanyParams.Count > 0)
                RemoveCompanyRows(entries, RequireSheet(sheetPaths, SheetCompany), rem.CompanyParams, shared);

            WriteAndValidate(xlsxPath, entries, delegate(string path)
            {
                Dictionary<string, List<XlsxRow>> book = ReopenAndCheck(path);
                HashSet<string> hdr = TenderHeaders(book);
                foreach (string col in rem.TenderColumns)
                    if (hdr.Contains(Util.Nfkc(col)))
                        throw new InvalidDataException("移除後欄位「" + col + "」仍存在，已還原原檔。");
                HashSet<string> keys = CompanyKeys(book);
                foreach (string p in rem.CompanyParams)
                    if (keys.Contains(Util.Nfkc(p)))
                        throw new InvalidDataException("移除後公司參數「" + p + "」仍存在，已還原原檔。");
                CheckRowCount(book, beforeRows);
            });
        }

        // ================= 標案清單：插入欄 =================
        static void InsertTenderColumns(List<KeyValuePair<string, byte[]>> entries,
            string wsPath, string tablePath, XlsxAdditions add, List<string> shared)
        {
            XDocument ws = LoadPart(entries, wsPath);
            XDocument tbl = tablePath != null ? LoadPart(entries, tablePath) : null;
            XElement sheetData = ws.Descendants(Main + "sheetData").First();
            XElement headerRow = sheetData.Elements(Main + "row").FirstOrDefault();
            if (headerRow == null) throw new InvalidDataException("「標案清單」缺少表頭列。");

            Dictionary<int, string> headers = HeaderMap(headerRow, shared);
            HashSet<string> existing = new HashSet<string>(headers.Values.Select(Util.Nfkc));
            if (tbl != null)
                foreach (XElement tc in tbl.Descendants(Main + "tableColumn"))
                    existing.Add(Util.Nfkc((string)tc.Attribute("name")));

            // 只加尚未存在者，保持傳入順序（服務層以 [產生:xxx, 參數...] 排好）
            List<string> toAdd = new List<string>();
            HashSet<string> seen = new HashSet<string>();
            foreach (string name in add.TenderColumns)
            {
                string nn = Util.Nfkc(name);
                if (existing.Contains(nn) || seen.Contains(nn)) continue;
                seen.Add(nn); toAdd.Add(name);
            }
            if (toAdd.Count == 0) { SavePart(entries, wsPath, ws); return; }

            int shift = toAdd.Count;
            int insertAt = ComputeInsertAt(headers);
            int lastRow = MaxRow(sheetData);
            string headerStyle = (string)headerRow.Elements(Main + "c").First().Attribute("s");

            // 1) 既有結構往右挪 shift 格
            ShiftCellRefs(sheetData, insertAt, shift);
            ShiftCols(ws, insertAt, shift);
            ShiftDataValidations(ws, insertAt, shift);

            // 2) 插入新表頭 cell（其餘資料列該欄留空）
            List<XElement> newValidations = new List<XElement>();
            for (int i = 0; i < toAdd.Count; i++)
            {
                int col = insertAt + i;
                string colL = ColLetter(col);
                XElement hc = new XElement(Main + "c", new XAttribute("r", colL + "1"),
                    new XElement(Main + "is", new XElement(Main + "t", toAdd[i])));
                hc.SetAttributeValue("t", "inlineStr");
                if (!string.IsNullOrEmpty(headerStyle)) hc.SetAttributeValue("s", headerStyle);
                headerRow.Add(hc);
                if (add.DropdownColumns.Contains(toAdd[i]))
                    newValidations.Add(MakeYesNoValidation(colL, lastRow));
            }
            SortRowCells(headerRow);

            // 3) dimension / spans
            int newLastCol = MaxCol(sheetData);
            SetDimension(ws, newLastCol, lastRow);
            if (headerRow.Attribute("spans") != null) headerRow.SetAttributeValue("spans", "1:" + newLastCol);

            // 4) 新產生欄的 是/否 下拉
            if (newValidations.Count > 0)
            {
                XElement dvs = EnsureDataValidations(ws);
                foreach (XElement dv in newValidations) dvs.Add(dv);
                dvs.SetAttributeValue("count", dvs.Elements(Main + "dataValidation").Count());
            }
            SavePart(entries, wsPath, ws);

            // 5) 同步 table：ref/autoFilter/tableColumns（在對應序位插入）
            if (tbl != null)
            {
                XElement tableEl = tbl.Root;
                string lastRowOfTable = Regex.Match((string)tableEl.Attribute("ref"), @":\D+(\d+)$").Groups[1].Value;
                string newRef = "A1:" + ColLetter(newLastCol) + lastRowOfTable;
                tableEl.SetAttributeValue("ref", newRef);
                XElement af = tableEl.Element(Main + "autoFilter");
                if (af != null) af.SetAttributeValue("ref", newRef);

                XElement cols = tableEl.Element(Main + "tableColumns");
                List<XElement> colEls = cols.Elements(Main + "tableColumn").ToList();
                int maxId = colEls.Select(c => IntAttr(c, "id")).DefaultIfEmpty(0).Max();
                // 於序位 (insertAt-1) 插入：全部插在「同一個錨點」前，依序即得正確順序
                int at = Math.Min(insertAt - 1, colEls.Count);
                XElement anchor = at < colEls.Count ? colEls[at] : null;
                for (int i = 0; i < toAdd.Count; i++)
                {
                    maxId++;
                    XElement tc = new XElement(Main + "tableColumn",
                        new XAttribute("id", maxId), new XAttribute("name", toAdd[i]));
                    if (anchor != null) anchor.AddBeforeSelf(tc);
                    else cols.Add(tc);
                }
                cols.SetAttributeValue("count", cols.Elements(Main + "tableColumn").Count());
                SavePart(entries, tablePath, tbl);
            }
        }

        // ================= 標案清單：移除欄 =================
        static void RemoveTenderColumns(List<KeyValuePair<string, byte[]>> entries,
            string wsPath, string tablePath, List<string> names, List<string> shared)
        {
            XDocument ws = LoadPart(entries, wsPath);
            XDocument tbl = tablePath != null ? LoadPart(entries, tablePath) : null;
            XElement sheetData = ws.Descendants(Main + "sheetData").First();
            XElement headerRow = sheetData.Elements(Main + "row").FirstOrDefault();
            if (headerRow == null) return;

            Dictionary<int, string> headers = HeaderMap(headerRow, shared);
            HashSet<string> want = new HashSet<string>(names.Select(Util.Nfkc));
            // 找出要移除的欄 index（由右至左）
            List<int> removeCols = headers.Where(kv => want.Contains(Util.Nfkc(kv.Value)))
                .Select(kv => kv.Key).OrderByDescending(x => x).ToList();
            if (removeCols.Count == 0) { SavePart(entries, wsPath, ws); return; }

            foreach (int col in removeCols)
            {
                // 刪各列該欄 cell
                foreach (XElement row in sheetData.Elements(Main + "row"))
                {
                    XElement c = row.Elements(Main + "c").FirstOrDefault(x => ColIndex((string)x.Attribute("r")) == col);
                    if (c != null) c.Remove();
                }
                // 欄 index > col 的往左挪 1
                ShiftCellRefs(sheetData, col + 1, -1);
                ShiftCols(ws, col + 1, -1);
                RemoveAndShiftDataValidations(ws, col);

                // table：移除對應 tableColumn（序位 col-1）、右界縮 1
                if (tbl != null)
                {
                    XElement cols = tbl.Root.Element(Main + "tableColumns");
                    List<XElement> colEls = cols.Elements(Main + "tableColumn").ToList();
                    if (col - 1 >= 0 && col - 1 < colEls.Count) colEls[col - 1].Remove();
                    cols.SetAttributeValue("count", cols.Elements(Main + "tableColumn").Count());
                }
            }

            int newLastCol = MaxCol(sheetData);
            int lastRow = MaxRow(sheetData);
            SetDimension(ws, newLastCol, lastRow);
            if (headerRow.Attribute("spans") != null) headerRow.SetAttributeValue("spans", "1:" + newLastCol);
            SavePart(entries, wsPath, ws);

            if (tbl != null)
            {
                string lastRowOfTable = Regex.Match((string)tbl.Root.Attribute("ref"), @":\D+(\d+)$").Groups[1].Value;
                string newRef = "A1:" + ColLetter(newLastCol) + lastRowOfTable;
                tbl.Root.SetAttributeValue("ref", newRef);
                XElement af = tbl.Root.Element(Main + "autoFilter");
                if (af != null) af.SetAttributeValue("ref", newRef);
                SavePart(entries, tablePath, tbl);
            }
        }

        // ================= 公司資料：加列 =================
        static void AppendCompanyRows(List<KeyValuePair<string, byte[]>> entries,
            string wsPath, List<string> paramNames, List<string> shared)
        {
            XDocument ws = LoadPart(entries, wsPath);
            XElement sheetData = ws.Descendants(Main + "sheetData").First();
            HashSet<string> existing = new HashSet<string>();
            int lastRow = 0;
            foreach (XElement r in sheetData.Elements(Main + "row"))
            {
                int rn; int.TryParse((string)r.Attribute("r"), out rn);
                if (rn > lastRow) lastRow = rn;
                XElement a = r.Elements(Main + "c").FirstOrDefault(c => ColIndex((string)c.Attribute("r")) == 1);
                if (a != null) { string t = Util.Nfkc(CellText(a, shared)); if (t != "") existing.Add(t); }
            }
            int added = 0;
            foreach (string name in paramNames)
            {
                if (existing.Contains(Util.Nfkc(name))) continue;
                lastRow++;
                XElement row = new XElement(Main + "row", new XAttribute("r", lastRow),
                    new XElement(Main + "c", new XAttribute("r", "A" + lastRow), new XAttribute("t", "inlineStr"),
                        new XElement(Main + "is", new XElement(Main + "t", name))),
                    new XElement(Main + "c", new XAttribute("r", "C" + lastRow), new XAttribute("t", "inlineStr"),
                        new XElement(Main + "is", new XElement(Main + "t", "（新範本需要，請填入固定值）"))));
                sheetData.Add(row);
                added++;
            }
            SetDimension(ws, 3, lastRow);
            SavePart(entries, wsPath, ws);
        }

        // ================= 公司資料：移除列 =================
        static void RemoveCompanyRows(List<KeyValuePair<string, byte[]>> entries,
            string wsPath, List<string> paramNames, List<string> shared)
        {
            XDocument ws = LoadPart(entries, wsPath);
            XElement sheetData = ws.Descendants(Main + "sheetData").First();
            HashSet<string> want = new HashSet<string>(paramNames.Select(Util.Nfkc));
            foreach (XElement row in sheetData.Elements(Main + "row").ToList())
            {
                XElement a = row.Elements(Main + "c").FirstOrDefault(c => ColIndex((string)c.Attribute("r")) == 1);
                if (a != null && want.Contains(Util.Nfkc(CellText(a, shared)))) row.Remove();
            }
            SetDimension(ws, 3, MaxRow(sheetData));
            SavePart(entries, wsPath, ws);
        }

        // ================= 位置/搬移 helper =================

        // 新欄插在「最後一個 產生: 欄」之後；找不到則插在「簽署日期」之後；再找不到 append 到最右。
        static int ComputeInsertAt(Dictionary<int, string> headers)
        {
            int lastGen = 0;
            foreach (KeyValuePair<int, string> kv in headers)
                if (Util.Nfkc(kv.Value).StartsWith("產生:") && kv.Key > lastGen) lastGen = kv.Key;
            if (lastGen > 0) return lastGen + 1;
            foreach (KeyValuePair<int, string> kv in headers)
                if (Util.Nfkc(kv.Value) == "簽署日期") return kv.Key + 1;
            return (headers.Count == 0 ? 1 : headers.Keys.Max() + 1);
        }

        static Dictionary<int, string> HeaderMap(XElement headerRow, List<string> shared)
        {
            Dictionary<int, string> map = new Dictionary<int, string>();
            foreach (XElement c in headerRow.Elements(Main + "c"))
            {
                int ci = ColIndex((string)c.Attribute("r"));
                string t = CellText(c, shared);
                if (ci > 0 && t != "") map[ci] = t;
            }
            return map;
        }

        // 把 sheetData 內欄index >= fromCol 的儲存格參照平移 delta（可負）。
        static void ShiftCellRefs(XElement sheetData, int fromCol, int delta)
        {
            foreach (XElement row in sheetData.Elements(Main + "row"))
                foreach (XElement c in row.Elements(Main + "c"))
                {
                    int ci = ColIndex((string)c.Attribute("r"));
                    if (ci >= fromCol) SetCellCol(c, ci + delta);
                }
        }

        static void SetCellCol(XElement c, int newCol)
        {
            string r = (string)c.Attribute("r");
            string digits = Regex.Match(r ?? "", @"\d+$").Value;
            c.SetAttributeValue("r", ColLetter(newCol) + digits);
        }

        static void SortRowCells(XElement row)
        {
            List<XElement> cells = row.Elements(Main + "c").OrderBy(c => ColIndex((string)c.Attribute("r"))).ToList();
            foreach (XElement c in cells) c.Remove();
            foreach (XElement c in cells) row.Add(c);
        }

        const int MaxSheetCol = 16384;   // Excel 欄數上限 XFD
        static void ShiftCols(XContainer ws, int fromCol, int delta)
        {
            XElement colsEl = ws.Descendants(Main + "cols").FirstOrDefault();
            if (colsEl == null) return;
            foreach (XElement col in colsEl.Elements(Main + "col").ToList())
            {
                int min = IntAttr(col, "min"), max = IntAttr(col, "max");
                if (min >= fromCol) min += delta;
                if (max >= fromCol) max += delta;
                if (min > MaxSheetCol) { col.Remove(); continue; } // 整段超出範圍 → 丟棄
                if (max > MaxSheetCol) max = MaxSheetCol;
                col.SetAttributeValue("min", min);
                col.SetAttributeValue("max", max);
            }
        }

        static void ShiftDataValidations(XContainer ws, int fromCol, int delta)
        {
            foreach (XElement dv in ws.Descendants(Main + "dataValidation"))
            {
                string sq = (string)dv.Attribute("sqref");
                if (sq == null) continue;
                dv.SetAttributeValue("sqref", ShiftSqref(sq, fromCol, delta));
            }
        }

        // 移除欄 col 時調整每段驗證範圍：整段在 col→丟棄；跨越 col→縮一欄（含「終點正好是 col」的情況）。
        static void RemoveAndShiftDataValidations(XContainer ws, int col)
        {
            XElement dvs = ws.Descendants(Main + "dataValidations").FirstOrDefault();
            if (dvs == null) return;
            foreach (XElement dv in dvs.Elements(Main + "dataValidation").ToList())
            {
                string sq = (string)dv.Attribute("sqref");
                if (sq == null) continue;
                List<string> keep = new List<string>();
                foreach (string range in sq.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    bool drop;
                    string adjusted = RemoveColFromRange(range, col, out drop);
                    if (!drop) keep.Add(adjusted);
                }
                if (keep.Count == 0) dv.Remove();
                else dv.SetAttributeValue("sqref", string.Join(" ", keep.ToArray()));
            }
            dvs.SetAttributeValue("count", dvs.Elements(Main + "dataValidation").Count());
            if (dvs.Elements(Main + "dataValidation").Count() == 0) dvs.Remove();
        }

        // 從範圍中扣掉欄 col。起點以 >col 才 −1（==col 時起點留在 col＝移入的鄰欄）；
        // 終點以 >=col 就 −1（涵蓋「終點正好是被移除欄」）。整段落空則 drop。
        static string RemoveColFromRange(string range, int col, out bool drop)
        {
            drop = false;
            int idx = range.IndexOf(':');
            string a = (idx < 0 ? range : range.Substring(0, idx)).Replace("$", "");
            string b = (idx < 0 ? range : range.Substring(idx + 1)).Replace("$", "");
            int c1 = ColIndex(a), c2 = ColIndex(b);
            if (c1 == col && c2 == col) { drop = true; return null; }
            int nc1 = c1 > col ? c1 - 1 : c1;
            int nc2 = c2 >= col ? c2 - 1 : c2;
            if (nc2 < nc1) { drop = true; return null; }
            string r1 = Regex.Match(a, @"\d+$").Value, r2 = Regex.Match(b, @"\d+$").Value;
            string sa = ColLetter(nc1) + r1;
            return idx < 0 ? sa : sa + ":" + ColLetter(nc2) + r2;
        }

        static string ShiftSqref(string sqref, int fromCol, int delta)
        {
            return string.Join(" ", sqref.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(r => ShiftRange(r, fromCol, delta)).ToArray());
        }

        static string ShiftRange(string range, int fromCol, int delta)
        {
            int idx = range.IndexOf(':');
            if (idx < 0) return ShiftRef(range, fromCol, delta);
            return ShiftRef(range.Substring(0, idx), fromCol, delta) + ":" +
                   ShiftRef(range.Substring(idx + 1), fromCol, delta);
        }

        static string ShiftRef(string cellRef, int fromCol, int delta)
        {
            cellRef = cellRef.Replace("$", "");
            int ci = ColIndex(cellRef);
            string digits = Regex.Match(cellRef, @"\d+$").Value;
            if (ci >= fromCol) ci += delta;
            return ColLetter(ci) + digits;
        }

        static void RangeCols(string range, out int c1, out int c2)
        {
            int idx = range.IndexOf(':');
            if (idx < 0) { c1 = c2 = ColIndex(range.Replace("$", "")); return; }
            c1 = ColIndex(range.Substring(0, idx).Replace("$", ""));
            c2 = ColIndex(range.Substring(idx + 1).Replace("$", ""));
        }

        static XElement MakeYesNoValidation(string colL, int lastRow)
        {
            string sq = lastRow >= 2 ? (colL + "2:" + colL + lastRow) : (colL + "2");
            return new XElement(Main + "dataValidation",
                new XAttribute("type", "list"), new XAttribute("allowBlank", "1"),
                new XAttribute("showInputMessage", "1"), new XAttribute("showErrorMessage", "1"),
                new XAttribute("sqref", sq), new XElement(Main + "formula1", "\"是,否\""));
        }

        static XElement EnsureDataValidations(XContainer ws)
        {
            XElement dvs = ws.Descendants(Main + "dataValidations").FirstOrDefault();
            if (dvs != null) return dvs;
            dvs = new XElement(Main + "dataValidations", new XAttribute("count", 0));
            XElement pm = ws.Descendants(Main + "pageMargins").FirstOrDefault();
            if (pm != null) pm.AddBeforeSelf(dvs);
            else ws.Descendants(Main + "sheetData").First().AddAfterSelf(dvs);
            return dvs;
        }

        static void SetDimension(XContainer ws, int lastCol, int lastRow)
        {
            XElement dim = ws.Descendants(Main + "dimension").FirstOrDefault();
            if (dim != null) dim.SetAttributeValue("ref", "A1:" + ColLetter(lastCol) + lastRow);
        }

        static int MaxCol(XElement sheetData)
        {
            int m = 1;
            foreach (XElement row in sheetData.Elements(Main + "row"))
                foreach (XElement c in row.Elements(Main + "c"))
                { int ci = ColIndex((string)c.Attribute("r")); if (ci > m) m = ci; }
            return m;
        }

        static int MaxRow(XElement sheetData)
        {
            int m = 1;
            foreach (XElement row in sheetData.Elements(Main + "row"))
            { int rn; if (int.TryParse((string)row.Attribute("r"), out rn) && rn > m) m = rn; }
            return m;
        }

        // ================= 寫回（備份→驗證→原子取代）=================
        static void WriteAndValidate(string xlsxPath, List<KeyValuePair<string, byte[]>> entries,
            Action<string> validatePath)
        {
            byte[] bytes = OdtWriter.Repack(entries);
            string backup = xlsxPath + ".備份-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".xlsx";
            string tmp = xlsxPath + ".tmp";
            File.Copy(xlsxPath, backup, true);
            try
            {
                File.WriteAllBytes(tmp, bytes);
                validatePath(tmp);   // 失敗丟例外
                try { File.Replace(tmp, xlsxPath, null); }
                catch (Exception ex)
                {
                    try { if (!File.Exists(xlsxPath) && File.Exists(backup)) File.Copy(backup, xlsxPath, true); } catch { }
                    if (ex is IOException || ex is UnauthorizedAccessException)
                        throw new IOException("無法更新 標案資料.xlsx —— 它可能正被 Excel 開啟，請關閉後再試一次。");
                    throw;
                }
                try { File.Delete(backup); } catch { }
            }
            catch
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
                try { if (File.Exists(xlsxPath)) File.Delete(backup); } catch { }
                throw;
            }
        }

        static Dictionary<string, List<XlsxRow>> ReopenAndCheck(string path)
        {
            try { return XlsxReader.Load(path); }
            catch (Exception ex) { throw new InvalidDataException("更新後的 Excel 無法讀取：" + ex.Message); }
        }

        static HashSet<string> TenderHeaders(Dictionary<string, List<XlsxRow>> book)
        {
            List<XlsxRow> rows = book.First(kv => Util.Nfkc(kv.Key) == SheetTenders).Value;
            return new HashSet<string>(rows[0].Cells.Select(c => Util.Nfkc(c)));
        }

        static HashSet<string> CompanyKeys(Dictionary<string, List<XlsxRow>> book)
        {
            List<XlsxRow> rows = book.First(kv => Util.Nfkc(kv.Key) == SheetCompany).Value;
            return new HashSet<string>(rows.Select(r => Util.Nfkc(r.Cell(0))));
        }

        static void CheckRowCount(Dictionary<string, List<XlsxRow>> book, int before)
        {
            if (before < 0) return;
            List<XlsxRow> rows = book.First(kv => Util.Nfkc(kv.Key) == SheetTenders).Value;
            int after = rows.Skip(1).Count(r => r.Cells.Any(c => (c ?? "").Trim() != ""));
            if (after != before)
                throw new InvalidDataException("更新後標案資料列數改變（" + before + "→" + after + "），已還原原檔。");
        }

        static int CountTenderDataRows(string xlsxPath)
        {
            try
            {
                Dictionary<string, List<XlsxRow>> book = XlsxReader.Load(xlsxPath);
                foreach (KeyValuePair<string, List<XlsxRow>> kv in book)
                    if (Util.Nfkc(kv.Key) == SheetTenders)
                        return kv.Value.Skip(1).Count(r => r.Cells.Any(c => (c ?? "").Trim() != ""));
            }
            catch { }
            return -1;
        }

        // ================= 解析 =================
        static string RequireSheet(Dictionary<string, string> sheetPaths, string name)
        {
            string p;
            if (!sheetPaths.TryGetValue(name, out p))
                throw new InvalidDataException("Excel 內找不到「" + name + "」工作表。");
            return p;
        }

        static Dictionary<string, string> ResolveSheetPaths(List<KeyValuePair<string, byte[]>> entries)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            XDocument wb = LoadPart(entries, "xl/workbook.xml");
            XDocument rels = LoadPartOrNull(entries, "xl/_rels/workbook.xml.rels");
            Dictionary<string, string> relTarget = new Dictionary<string, string>();
            if (rels != null)
                foreach (XElement rel in rels.Descendants().Where(e => e.Name.LocalName == "Relationship"))
                {
                    string id = (string)rel.Attribute("Id"), target = (string)rel.Attribute("Target");
                    if (id != null && target != null) relTarget[id] = ResolveXlPath(target);
                }
            int i = 0;
            foreach (XElement sheet in wb.Descendants().Where(e => e.Name.LocalName == "sheet"))
            {
                i++;
                string name = (string)sheet.Attribute("name");
                string rid = sheet.Attributes().Where(a => a.Name.LocalName == "id").Select(a => a.Value).FirstOrDefault();
                string path = (rid != null && relTarget.ContainsKey(rid)) ? relTarget[rid] : ("xl/worksheets/sheet" + i + ".xml");
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
            XDocument rels = LoadPartOrNull(entries, "xl/worksheets/_rels/" + Path.GetFileName(wsPath) + ".rels");
            if (rels == null) return null;
            foreach (XElement rel in rels.Descendants().Where(e => e.Name.LocalName == "Relationship"))
                if ((string)rel.Attribute("Id") == rid)
                    return NormalizeRelative(wsPath, (string)rel.Attribute("Target"));
            return null;
        }

        // ================= zip / XML part 存取 =================
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
                s.Encoding = new UTF8Encoding(false); s.Indent = false;
                using (XmlWriter xw = XmlWriter.Create(ms, s)) { doc.Save(xw); }
                bytes = ms.ToArray();
            }
            entries[idx] = new KeyValuePair<string, byte[]>(entries[idx].Key, bytes);
        }

        // ================= 小工具 =================
        static string CellText(XElement c, List<string> shared)
        {
            string t = (string)c.Attribute("t");
            if (t == "inlineStr") return string.Concat(c.Descendants(Main + "t").Select(x => x.Value));
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

        static int IntAttr(XElement el, string name)
        {
            int v; return int.TryParse((string)el.Attribute(name), out v) ? v : 0;
        }

        static int ColIndex(string cellRef)
        {
            if (cellRef == null) return 0;
            int col = 0;
            foreach (char ch in cellRef)
            {
                char u = char.ToUpperInvariant(ch);
                if (u >= 'A' && u <= 'Z') col = col * 26 + (u - 'A' + 1);
                else if (char.IsDigit(ch) || ch == '$') { if (col > 0) break; }
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
            relative = relative.Replace('\\', '/');
            if (relative.StartsWith("/")) return relative.TrimStart('/');
            List<string> parts = basePartPath.Split('/').ToList();
            parts.RemoveAt(parts.Count - 1);
            foreach (string seg in relative.Split('/'))
            {
                if (seg == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); }
                else if (seg != "." && seg != "") parts.Add(seg);
            }
            return string.Join("/", parts);
        }
    }
}
