using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace TenderDocGen
{
    /// <summary>一列標案在產生前的完整計畫（參數已解析、錯誤已檢查）。</summary>
    class RowPlan
    {
        public int RowNumber;                        // Excel 實際列號
        public string TenderName = "";               // 標案名稱（原文）
        public string FolderName;                    // 清理後資料夾名（null = 無法建立）
        public List<TemplateInfo> Docs = new List<TemplateInfo>();
        public Dictionary<string, string> Values = new Dictionary<string, string>();
        public List<string> Errors = new List<string>();
        public List<string> Warnings = new List<string>();
        public string ResolvedDate = "";             // 顯示用：實際帶入的民國日期
    }

    /// <summary>整份 Excel 解析後的產生計畫。</summary>
    class GenerationPlan
    {
        public List<RowPlan> Rows = new List<RowPlan>();
        public Dictionary<string, string> CompanyDefaults =
            new Dictionary<string, string>();
        public bool KeepRed;                         // 替換後維持紅色（校對用），預設黑色
        public List<string> GlobalErrors = new List<string>();
    }

    /// <summary>一列的產生結果。</summary>
    class RowResult
    {
        public RowPlan Plan;
        public List<string> Generated = new List<string>();  // 產生的檔案
        public List<string> Skipped = new List<string>();    // 已存在而略過的檔案
        public List<string> Errors = new List<string>();
    }

    static class Planner
    {
        public const string SheetTenders = "標案清單";
        public const string SheetCompany = "公司資料";
        public const string ColTenderName = "標案名稱";
        public const string ColSignDate = "簽署日期";
        public const string GeneratePrefix = "產生:";
        public const string KeyColor = "替換後文字顏色";
        public const int MaxFolderNameLength = 60;

        /// <summary>讀 Excel＋範本庫 → 產生計畫。所有可預見的問題都收進 Errors/GlobalErrors，不丟例外。</summary>
        public static GenerationPlan BuildPlan(string xlsxPath, TemplateStore store)
        {
            GenerationPlan plan = new GenerationPlan();
            Dictionary<string, List<XlsxRow>> book = XlsxReader.Load(xlsxPath);

            // --- 公司資料（鍵值表）---
            List<XlsxRow> companyRows = FindSheet(book, SheetCompany);
            if (companyRows != null)
            {
                foreach (XlsxRow r in companyRows)
                {
                    string key = Util.Nfkc(r.Cell(0));
                    string val = r.Cell(1).Trim();
                    if (key == "" || key == "參數名稱" || key == "項目") continue;
                    plan.CompanyDefaults[key] = val;
                }
            }
            string colorSetting;
            if (plan.CompanyDefaults.TryGetValue(KeyColor, out colorSetting))
                plan.KeepRed = Util.Nfkc(colorSetting).Contains("紅");

            // --- 標案清單 ---
            List<XlsxRow> rows = FindSheet(book, SheetTenders);
            if (rows == null)
            {
                plan.GlobalErrors.Add("Excel 裡找不到工作表「" + SheetTenders + "」。");
                return plan;
            }
            if (rows.Count == 0)
            {
                plan.GlobalErrors.Add("工作表「" + SheetTenders + "」是空的。");
                return plan;
            }

            // 表頭對應（NFKC 正規化後比對；容忍全形冒號、多餘空白）
            XlsxRow header = rows[0];
            Dictionary<string, int> colIndex = new Dictionary<string, int>();
            for (int i = 0; i < header.Cells.Count; i++)
            {
                string h = Util.Nfkc(header.Cell(i));
                if (h != "" && !colIndex.ContainsKey(h)) colIndex[h] = i;
            }
            if (!colIndex.ContainsKey(ColTenderName))
            {
                plan.GlobalErrors.Add("工作表「" + SheetTenders + "」第一列缺少「" + ColTenderName + "」欄。");
                return plan;
            }

            // 範本本身無法解析（檔案損壞等）→ 全域錯誤
            foreach (TemplateInfo tpl in store.Templates)
                if (tpl.LoadError != null)
                    plan.GlobalErrors.Add("範本「" + tpl.FileName + "」無法讀取：" + tpl.LoadError);

            // 產生:xxx 欄與範本檔的對應檢查（兩邊落單都提醒，不靜默）
            Dictionary<string, int> genCols = new Dictionary<string, int>();
            foreach (KeyValuePair<string, int> kv in colIndex)
            {
                if (!kv.Key.StartsWith(GeneratePrefix)) continue;
                string baseName = kv.Key.Substring(GeneratePrefix.Length).Trim();
                TemplateInfo tpl = store.Templates.FirstOrDefault(t => Util.Nfkc(t.BaseName) == baseName);
                if (tpl == null)
                    plan.GlobalErrors.Add("Excel 欄「" + kv.Key + "」找不到對應的範本檔，請確認 範本 資料夾。");
                else
                    genCols[tpl.FileName] = kv.Value;
            }

            // --- 逐列 ---
            HashSet<string> seenFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<int> genColIdx = new HashSet<int>(genCols.Values);
            for (int ri = 1; ri < rows.Count; ri++)
            {
                XlsxRow row = rows[ri];
                bool allEmpty = row.Cells.All(c => (c ?? "").Trim() == "");
                if (allEmpty) continue;

                RowPlan rp = new RowPlan();
                rp.RowNumber = row.RowNumber;
                rp.TenderName = row.Cell(colIndex[ColTenderName]).Trim();
                if (rp.TenderName == "")
                {
                    // 只有「產生:」欄有值（表格新列的下拉預設）視同空白列，靜默跳過
                    bool onlyGenCols = true;
                    for (int ci = 0; ci < row.Cells.Count; ci++)
                        if (!genColIdx.Contains(ci) && row.Cell(ci).Trim() != "") { onlyGenCols = false; break; }
                    if (onlyGenCols) continue;

                    rp.Errors.Add("缺少「標案名稱」。");
                    plan.Rows.Add(rp);
                    continue;
                }

                rp.FolderName = Util.SanitizeFolderName(rp.TenderName, MaxFolderNameLength, rp.Warnings);
                if (rp.FolderName == null)
                    rp.Errors.Add("標案名稱清理後為空，無法建立資料夾。");
                else if (!seenFolders.Add(rp.FolderName))
                    rp.Errors.Add("與其他列的標案名稱重複（資料夾名相同會互相覆蓋）。");

                // 這一列要產生哪些文件：有「產生:xxx」欄就看欄值；範本沒有對應欄 → 預設產生
                foreach (TemplateInfo tpl in store.Templates)
                {
                    int gc;
                    bool generate = true;
                    if (genCols.TryGetValue(tpl.FileName, out gc))
                        generate = Util.IsTruthy(row.Cell(gc));
                    if (generate) rp.Docs.Add(tpl);
                }
                if (rp.Docs.Count == 0)
                    rp.Warnings.Add("所有文件都設為「否」，此列不會產生任何檔案。");

                // 參數解析：列值 → 公司資料預設 → 特殊 fallback
                // 所需參數以「範本內實際的 ${token}」為準（不依賴 tokens.txt，缺 manifest 仍正確）
                HashSet<string> needed = new HashSet<string>();
                foreach (TemplateInfo tpl in rp.Docs)
                    foreach (string tk in tpl.Tokens) needed.Add(tk);

                ResolveDate(rp, row, colIndex, needed);

                foreach (string token in needed)
                {
                    if (rp.Values.ContainsKey(token)) continue;   // 日期已處理
                    string val = "";
                    int ci;
                    if (colIndex.TryGetValue(token, out ci)) val = row.Cell(ci).Trim();
                    if (val == "" && plan.CompanyDefaults.ContainsKey(token))
                        val = plan.CompanyDefaults[token].Trim();
                    if (val == "" && token == "案件名稱")
                        val = rp.TenderName;                      // 案件名稱缺省 = 標案名稱
                    if (val == "")
                        rp.Errors.Add("缺少參數「" + token + "」：請填在標案清單的「" + token +
                                      "」欄或「公司資料」工作表。");
                    else
                        rp.Values[token] = val;
                }

                plan.Rows.Add(rp);
            }
            if (plan.Rows.Count == 0)
                plan.GlobalErrors.Add("工作表「" + SheetTenders + "」沒有任何資料列。");
            return plan;
        }

        /// <summary>簽署日期欄 → 簽署年/月/日。缺省用今天（民國），並記錄實際帶入值供畫面顯示。</summary>
        static void ResolveDate(RowPlan rp, XlsxRow row, Dictionary<string, int> colIndex, HashSet<string> needed)
        {
            if (!needed.Contains("簽署年") && !needed.Contains("簽署月") && !needed.Contains("簽署日")) return;

            int y, m, d;
            int ci;
            string raw = "";
            if (colIndex.TryGetValue(ColSignDate, out ci)) raw = row.Cell(ci).Trim();

            if (raw == "")
            {
                Util.TodayRoc(out y, out m, out d);
                rp.Warnings.Add("未填「簽署日期」，已帶入今天：" + y + "年" + m + "月" + d + "日");
            }
            else if (!Util.TryParseRocDate(raw, out y, out m, out d))
            {
                rp.Errors.Add("「簽署日期」格式無法解析：「" + raw + "」。請用民國格式，例如 115年1月1日 或 115/1/1。");
                return;
            }
            rp.ResolvedDate = y + "年" + m + "月" + d + "日";
            rp.Values["簽署年"] = y.ToString();
            rp.Values["簽署月"] = m.ToString();
            rp.Values["簽署日"] = d.ToString();
        }

        static List<XlsxRow> FindSheet(Dictionary<string, List<XlsxRow>> book, string name)
        {
            foreach (KeyValuePair<string, List<XlsxRow>> kv in book)
                if (Util.Nfkc(kv.Key) == name) return kv.Value;
            return null;
        }
    }

    static class Generator
    {
        /// <summary>執行一列的產生。plan.Errors 非空時直接回報不產生。</summary>
        public static RowResult GenerateRow(RowPlan rp, string outputRoot, bool overwrite, bool keepRed)
        {
            RowResult result = new RowResult();
            result.Plan = rp;
            if (rp.Errors.Count > 0)
            {
                result.Errors.AddRange(rp.Errors);
                return result;
            }

            string folder = Path.Combine(outputRoot, rp.FolderName);
            try
            {
                Directory.CreateDirectory(folder);
            }
            catch (Exception ex)
            {
                result.Errors.Add("無法建立資料夾「" + folder + "」：" + ex.Message);
                return result;
            }

            foreach (TemplateInfo tpl in rp.Docs)
            {
                string outPath = Path.Combine(folder, tpl.FileName);
                if (File.Exists(outPath) && !overwrite)
                {
                    result.Skipped.Add(tpl.FileName);
                    continue;
                }
                string tmpPath = outPath + ".tmp";
                try
                {
                    HashSet<string> usedTokens;
                    byte[] bytes = OdtWriter.Generate(tpl, rp.Values, keepRed, out usedTokens);

                    // 先寫暫存檔再改名，避免中斷留下半成品
                    File.WriteAllBytes(tmpPath, bytes);

                    // 只驗證「此範本實際用到」的參數值有出現在文件中
                    Dictionary<string, string> expectHere = new Dictionary<string, string>();
                    foreach (string tk in usedTokens)
                        if (rp.Values.ContainsKey(tk)) expectHere[tk] = rp.Values[tk];
                    List<string> problems = Validator.ValidateOdt(tmpPath, expectHere, tpl.Tokens);
                    if (problems.Count > 0)
                    {
                        result.Errors.Add(tpl.FileName + " 產出驗證失敗：" + string.Join("；", problems));
                        continue;
                    }
                    if (File.Exists(outPath)) File.Delete(outPath);
                    File.Move(tmpPath, outPath);
                    result.Generated.Add(tpl.FileName);
                }
                catch (IOException ex)
                {
                    // 目標檔可能正被 Word/LibreOffice 開著而無法刪除/取代
                    result.Errors.Add(tpl.FileName + "：無法寫入，該檔可能正被開啟中，請關閉後再試。（" + ex.Message + "）");
                }
                catch (UnauthorizedAccessException ex)
                {
                    result.Errors.Add(tpl.FileName + "：沒有寫入權限（" + ex.Message + "）");
                }
                catch (Exception ex)
                {
                    result.Errors.Add(tpl.FileName + "：" + ex.Message);
                }
                finally
                {
                    // 不論成功失敗都清掉殘留的暫存檔（例如覆寫時目標被鎖住）
                    try { if (File.Exists(tmpPath)) File.Delete(tmpPath); }
                    catch { /* 暫存檔清不掉不影響結果，忽略 */ }
                }
            }
            return result;
        }
    }
}
