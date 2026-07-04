using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace TenderDocGen
{
    class AddTemplateResult
    {
        public bool Ok;
        public string TemplateFile;                          // 存好的範本檔名
        public List<string> AddedTenderColumns = new List<string>();
        public List<string> AddedCompanyParams = new List<string>();
        public string Error;                                 // null = 成功
    }

    /// <summary>移除範本的預覽/結果。</summary>
    class RemovalPlan
    {
        public bool Ok;
        public string TemplateFile;
        public List<string> TenderColumns = new List<string>();   // 將移除的標案清單欄（含 產生:xxx）
        public List<string> CompanyParams = new List<string>();   // 將移除的公司資料列
        public string Error;                                      // null = 可執行/成功
    }

    /// <summary>
    /// 「新增範本」的無 UI 核心邏輯（可被 selftest 直接呼叫）：
    /// 正規化 ODT → 更新 Excel（加欄/加列）→ 存範本 → 重寫 tokens.txt。
    /// 順序為「先更新 Excel、再存範本」，Excel 更新有自身備份/還原；若 Excel 更新失敗
    /// 則不存範本，狀態保持乾淨。
    /// </summary>
    static class AddTemplateService
    {
        // 開發時就存在的 9 個參數（沿用既有 Excel 欄位/公司資料，不需再加）
        public static readonly string[] ExistingTokens =
        {
            "廠商名稱", "負責人姓名", "聯絡電話", "廠商地址",
            "機關名稱", "案件名稱", "簽署年", "簽署月", "簽署日"
        };
        public static readonly HashSet<string> CompanyTokens = new HashSet<string>
        {
            "廠商名稱", "負責人姓名", "聯絡電話", "廠商地址"
        };

        public static bool IsExisting(string token) { return ExistingTokens.Contains(token); }

        public static string TemplateDir(string baseDir) { return Path.Combine(baseDir, "範本"); }
        public static string XlsxPath(string baseDir) { return Path.Combine(baseDir, "標案資料.xlsx"); }

        /// <summary>
        /// runToToken：runId → token（不含或空 = 略過）。
        /// newTokenIsCompany：使用者自建的新 token → 是否為公司固定資料。
        /// </summary>
        public static AddTemplateResult Commit(string odtPath, string templateName,
            Dictionary<string, string> runToToken, Dictionary<string, bool> newTokenIsCompany,
            string baseDir, bool overwrite)
        {
            AddTemplateResult result = new AddTemplateResult();
            try
            {
                string name = SanitizeTemplateName(templateName);
                if (name == "") { result.Error = "請輸入有效的範本名稱。"; return result; }

                string templateDir = TemplateDir(baseDir);
                if (!Directory.Exists(templateDir)) Directory.CreateDirectory(templateDir);
                string targetPath = Path.Combine(templateDir, name + ".odt");
                if (File.Exists(targetPath) && !overwrite)
                { result.Error = "已存在同名範本「" + name + "」。"; return result; }

                // 至少要對應一個參數
                if (runToToken == null || runToToken.Values.All(v => string.IsNullOrEmpty(v)))
                { result.Error = "尚未指定任何參數對應，無法建立範本。"; return result; }

                // 參數名稱不可含會破壞 ${token} 或 XML 的字元
                foreach (string tk in runToToken.Values.Where(v => !string.IsNullOrEmpty(v)).Distinct())
                    if (!IsValidTokenName(tk))
                    { result.Error = "參數名稱「" + tk + "」含有不允許的字元（不可包含 { } $ 或控制字元）。"; return result; }

                // 正規化
                TemplateBuilder tb = TemplateBuilder.Load(odtPath);
                tb.ExtractRuns();
                byte[] bytes = tb.Build(runToToken);
                List<string> problems = TemplateBuilder.ValidateTemplateBytes(bytes);
                if (problems.Count > 0)
                { result.Error = "正規化產出無效：" + string.Join("；", problems); return result; }

                // 依實際用到的 token 決定 Excel 要加什麼
                HashSet<string> usedTokens = new HashSet<string>(
                    runToToken.Values.Where(v => !string.IsNullOrEmpty(v)));
                XlsxAdditions add = new XlsxAdditions();
                List<string> newPerCase = new List<string>();
                foreach (string token in usedTokens)
                {
                    if (IsExisting(token)) continue;              // 沿用既有欄位
                    bool isCompany;
                    newTokenIsCompany.TryGetValue(token, out isCompany);
                    if (isCompany) { if (!add.CompanyParams.Contains(token)) add.CompanyParams.Add(token); }
                    else { if (!newPerCase.Contains(token)) newPerCase.Add(token); }
                }
                // 順序：先「產生:<範本名>」（讓產生欄群組連續），再新增的每案參數欄
                string genCol = "產生:" + name;
                add.TenderColumns.Add(genCol);
                add.DropdownColumns.Add(genCol);
                add.TenderColumns.AddRange(newPerCase);

                // 先更新 Excel（有自身備份/還原；失敗就不存範本）
                string xlsx = XlsxPath(baseDir);
                if (File.Exists(xlsx))
                    XlsxEditor.Apply(xlsx, add);

                // 存範本（暫存→改名，避免半成品）
                string tmp = targetPath + ".tmp";
                File.WriteAllBytes(tmp, bytes);
                if (File.Exists(targetPath)) File.Delete(targetPath);
                File.Move(tmp, targetPath);

                // 重寫 tokens.txt（可重建的 manifest，失敗不致命，不讓它把已完成的建立判為失敗）
                try { TemplateStore.RewriteTokensFile(templateDir); } catch { }

                result.Ok = true;
                result.TemplateFile = name + ".odt";
                result.AddedTenderColumns = add.TenderColumns;
                result.AddedCompanyParams = add.CompanyParams;
                return result;
            }
            catch (Exception ex)
            {
                result.Error = ex.Message;
                return result;
            }
        }

        /// <summary>移除範本的預覽/結果：將移除（或已移除）哪些 Excel 欄位與參數列。</summary>
        public static RemovalPlan PreviewRemoval(string baseName, string baseDir)
        {
            RemovalPlan plan = new RemovalPlan();
            plan.TemplateFile = baseName + ".odt";
            string templateDir = TemplateDir(baseDir);
            string target = Path.Combine(templateDir, baseName + ".odt");
            if (!File.Exists(target)) { plan.Error = "找不到範本「" + baseName + "」。"; return plan; }

            List<string> allFiles = Directory.GetFiles(templateDir, "*.odt")
                .Where(f => { string n = Path.GetFileName(f); return !n.StartsWith("~") && !n.StartsWith("."); })
                .ToList();
            if (allFiles.Count <= 1) { plan.Error = "至少需保留一份範本，無法移除最後一份。"; return plan; }

            HashSet<string> targetTokens;
            try { targetTokens = OdtWriter.ExtractTokens(target); }
            catch (Exception ex) { plan.Error = "無法讀取範本：" + ex.Message; return plan; }

            // 其他範本用到的所有 token（共用者不可移除）
            HashSet<string> otherTokens = new HashSet<string>();
            foreach (string f in allFiles)
            {
                if (string.Equals(f, target, StringComparison.OrdinalIgnoreCase)) continue;
                try { foreach (string tk in OdtWriter.ExtractTokens(f)) otherTokens.Add(tk); }
                catch { }
            }

            // 依 token「目前實際位在 Excel 哪個工作表」判斷是公司資料列還是標案清單欄
            // （不能只靠內建 CompanyTokens——使用者自建的公司固定參數不在其中）。
            HashSet<string> companyKeys = new HashSet<string>();
            HashSet<string> tenderHeaders = new HashSet<string>();
            try
            {
                string xlsx = XlsxPath(baseDir);
                if (File.Exists(xlsx))
                {
                    Dictionary<string, List<XlsxRow>> book = XlsxReader.Load(xlsx);
                    foreach (KeyValuePair<string, List<XlsxRow>> kv in book)
                    {
                        string sn = Util.Nfkc(kv.Key);
                        if (sn == "公司資料")
                            foreach (XlsxRow r in kv.Value) companyKeys.Add(Util.Nfkc(r.Cell(0)));
                        else if (sn == "標案清單" && kv.Value.Count > 0)
                            foreach (string h in kv.Value[0].Cells) tenderHeaders.Add(Util.Nfkc(h));
                    }
                }
            }
            catch { }

            foreach (string token in targetTokens)
            {
                if (IsExisting(token) || otherTokens.Contains(token)) continue; // 內建或共用 → 保留
                string nn = Util.Nfkc(token);
                if (companyKeys.Contains(nn)) plan.CompanyParams.Add(token);
                else if (tenderHeaders.Contains(nn)) plan.TenderColumns.Add(token);
                else if (CompanyTokens.Contains(token)) plan.CompanyParams.Add(token);  // 後備
                else plan.TenderColumns.Add(token);
            }
            plan.TenderColumns.Add("產生:" + baseName);   // 該範本的產生欄一律移除
            return plan;
        }

        /// <summary>執行移除：先更新 Excel（移欄/列）、再刪 odt、重寫 tokens.txt。</summary>
        public static RemovalPlan RemoveTemplate(string baseName, string baseDir)
        {
            RemovalPlan plan = PreviewRemoval(baseName, baseDir);
            if (plan.Error != null) return plan;
            try
            {
                XlsxRemovals rem = new XlsxRemovals();
                rem.TenderColumns = plan.TenderColumns;
                rem.CompanyParams = plan.CompanyParams;
                string xlsx = XlsxPath(baseDir);
                if (File.Exists(xlsx)) XlsxEditor.Remove(xlsx, rem);   // Excel 先移除（有備份/還原）

                // 刪範本檔（可能被其他程式暫時鎖住 → 短暫重試）
                string odt = Path.Combine(TemplateDir(baseDir), baseName + ".odt");
                Exception delErr = null;
                for (int i = 0; i < 3; i++)
                {
                    try { if (File.Exists(odt)) File.Delete(odt); delErr = null; break; }
                    catch (Exception ex) { delErr = ex; System.Threading.Thread.Sleep(250); }
                }
                if (delErr != null)
                {
                    // Excel 已更新，但檔案刪不掉：明確告知，重試可完成（Excel 部分為冪等）
                    plan.Error = "Excel 欄位已移除，但範本檔「" + baseName + ".odt」刪除失敗" +
                                 "（可能被其他程式開啟）。請關閉後於「範本管理」再按一次移除即可完成。";
                    return plan;
                }
                try { TemplateStore.RewriteTokensFile(TemplateDir(baseDir)); } catch { }

                plan.Ok = true;
                return plan;
            }
            catch (Exception ex) { plan.Error = ex.Message; return plan; }
        }

        /// <summary>參數名稱是否安全：非空、不含 { } $ 與控制字元（會破壞 ${token} 或 XML）。</summary>
        public static bool IsValidTokenName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            foreach (char c in name)
                if (c == '{' || c == '}' || c == '$' || char.IsControl(c)) return false;
            return true;
        }

        static string SanitizeTemplateName(string name)
        {
            if (name == null) return "";
            List<string> warn = new List<string>();
            string s = Util.SanitizeFolderName(name.Trim(), 80, warn);   // 沿用資料夾名清理規則
            return s ?? "";
        }
    }
}
