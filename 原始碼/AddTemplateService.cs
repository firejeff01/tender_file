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
                foreach (string token in usedTokens)
                {
                    if (IsExisting(token)) continue;              // 沿用既有欄位
                    bool isCompany;
                    newTokenIsCompany.TryGetValue(token, out isCompany);
                    if (isCompany) { if (!add.CompanyParams.Contains(token)) add.CompanyParams.Add(token); }
                    else { if (!add.TenderColumns.Contains(token)) add.TenderColumns.Add(token); }
                }
                // 一律加「產生:<範本名>」欄（是/否下拉）
                string genCol = "產生:" + name;
                add.TenderColumns.Add(genCol);
                add.DropdownColumns.Add(genCol);

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
