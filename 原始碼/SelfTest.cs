using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace TenderDocGen
{
    /// <summary>
    /// 開發用自我測試（/selftest）。
    /// 涵蓋：Util 單元測試、xlsx 解析邊角（inlineStr、缺 r 屬性、科學記號、布林）、
    /// 端到端產生（跳脫、換行、超長/非法/保留資料夾名、重複列、略過與覆寫、紅黑設定、
    /// 缺參數錯誤、範本竄改偵測），並讀取真實的 標案資料.xlsx 驗證表頭。
    /// 結果寫入 selftest.log，exit code = 失敗數。
    /// </summary>
    static class SelfTest
    {
        static readonly List<string> Log = new List<string>();
        static int _failures;

        static void Check(bool cond, string name, string detail)
        {
            if (cond) Log.Add("OK   " + name);
            else { _failures++; Log.Add("FAIL " + name + (detail != null ? "：" + detail : "")); }
        }

        public static int Run(string[] args)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string tmpDir = Path.Combine(baseDir, "selftest_tmp");
            try
            {
                if (Directory.Exists(tmpDir)) Directory.Delete(tmpDir, true);
                Directory.CreateDirectory(tmpDir);

                TestUtil();
                TestSettings(tmpDir);
                TemplateStore store = TestTemplateStore(baseDir);
                TestRealXlsx(baseDir, store);
                if (store != null)
                {
                    TestEndToEnd(store, tmpDir);
                    TestKeepRed(store, tmpDir);
                    TestMissingParam(store, tmpDir);
                    TestTamperDetection(store, tmpDir);
                    TestTokensTxtMissing(baseDir, tmpDir);
                    TestLiteralDollarBrace(store, tmpDir);
                    TestPhoneticRuby();
                    TestRegenerateClearsFolder(store, tmpDir);
                    TestTemplateBuilder(tmpDir);
                    TestDocxToOdt(tmpDir);
                    TestMergeWhitespaceFix(tmpDir);
                    TestXlsxEditor(baseDir, tmpDir);
                    TestXlsxInsertPositionAndRemove(baseDir, tmpDir);
                    TestValidationShrinkOnRemove(baseDir, tmpDir);
                    TestAddTemplateService(baseDir, tmpDir);
                    TestRemoveTemplateService(baseDir, tmpDir);
                }
            }
            catch (Exception ex)
            {
                _failures++;
                Log.Add("FAIL 未預期例外：" + ex);
            }

            Log.Add("");
            Log.Add(string.Format("結果：{0} 項失敗（共 {1} 項檢查）", _failures, Log.Count - 2));
            File.WriteAllLines(Path.Combine(baseDir, "selftest.log"), Log, new UTF8Encoding(true));
            return _failures;
        }

        // ============================== Util ==============================
        static void TestUtil()
        {
            int y, m, d;
            Check(Util.TryParseRocDate("115年1月1日", out y, out m, out d) && y == 115 && m == 1 && d == 1,
                "日期解析 115年1月1日", null);
            Check(Util.TryParseRocDate("116/2/28", out y, out m, out d) && y == 116 && m == 2 && d == 28,
                "日期解析 116/2/28", null);
            Check(Util.TryParseRocDate("１１５年３月５日", out y, out m, out d) && y == 115 && m == 3 && d == 5,
                "日期解析 全形數字", null);
            Check(!Util.TryParseRocDate("115年13月1日", out y, out m, out d), "日期解析 拒絕 13 月", null);
            Check(!Util.TryParseRocDate("115年2月30日", out y, out m, out d), "日期解析 拒絕 2月30日", null);
            Check(!Util.TryParseRocDate("116年4月31日", out y, out m, out d), "日期解析 拒絕 4月31日", null);
            Check(!Util.TryParseRocDate("115年2月29日", out y, out m, out d), "日期解析 拒絕 非閏年2月29日", null);
            Check(Util.TryParseRocDate("113年2月29日", out y, out m, out d) && m == 2 && d == 29,
                "日期解析 接受 閏年2月29日（民國113=2024）", null);
            Check(Util.TryParseRocDate("45658", out y, out m, out d) &&
                  y == DateTime.FromOADate(45658).Year - 1911, "日期解析 Excel 序列值", null);

            List<string> warn = new List<string>();
            string s1 = Util.SanitizeFolderName("A/B\\C:D*E?F\"G<H>I|J", 60, warn);
            Check(s1 == "A／B＼C：D＊E？F＂G＜H＞I｜J", "資料夾名 非法字元轉全形", s1);
            string s2 = Util.SanitizeFolderName("CON", 60, warn);
            Check(s2 == "＿CON", "資料夾名 保留裝置名", s2);
            string s3 = Util.SanitizeFolderName("測試名稱...  ", 60, warn);
            Check(s3 == "測試名稱", "資料夾名 去結尾點與空白", s3);
            string s4 = Util.SanitizeFolderName(new string('長', 80), 60, warn);
            Check(s4 != null && s4.Length == 60, "資料夾名 截長至 60", s4 != null ? s4.Length.ToString() : "null");
            string s5 = Util.SanitizeFolderName(new string('.', 60) + "案", 60, warn);
            Check(s5 == null, "資料夾名 截長後全空→null（不寫進根目錄）", s5 == null ? "null" : "'" + s5 + "'");

            Check(Util.ExpandScientific("9.12345678E8") == "912345678", "科學記號展開", null);
            Check(Util.IsTruthy("") && Util.IsTruthy("是") && !Util.IsTruthy("否") && !Util.IsTruthy(" x "),
                "是/否判定", null);
            Check(Util.Nfkc("產生：保密切結書") == "產生:保密切結書", "NFKC 全形冒號", null);
        }

        // ============================== 設定檔 ==============================
        static void TestSettings(string tmpDir)
        {
            string dir = Path.Combine(tmpDir, "settings");
            Directory.CreateDirectory(dir);
            Settings s1 = new Settings(dir);
            Check(s1.Get(Settings.KeyOutputDir, "預設") == "預設", "設定 缺檔用 fallback", null);
            s1.Set(Settings.KeyOutputDir, @"D:\共用\標案輸出");
            Check(File.Exists(Path.Combine(dir, "設定.txt")), "設定 已寫出 設定.txt", null);
            // 以新實例重讀，模擬下次啟動
            Settings s2 = new Settings(dir);
            Check(s2.Get(Settings.KeyOutputDir, "") == @"D:\共用\標案輸出", "設定 重讀持久化值", null);
            // 含全形字與空白的路徑
            s2.Set(Settings.KeyOutputDir, @"\\伺服器\共用 資料夾\輸出");
            Settings s3 = new Settings(dir);
            Check(s3.Get(Settings.KeyOutputDir, "") == @"\\伺服器\共用 資料夾\輸出", "設定 UNC/全形路徑", null);
        }

        // ============================== 範本庫 ==============================
        static TemplateStore TestTemplateStore(string baseDir)
        {
            TemplateStore store = null;
            try
            {
                store = TemplateStore.Load(Path.Combine(baseDir, "範本"));
                Check(store.Templates.Count == 3, "範本庫 3 個範本", store.Templates.Count.ToString());
                TemplateInfo secrecy = store.Templates.FirstOrDefault(t => t.BaseName == "保密同意書");
                Check(secrecy != null && secrecy.ExpectedTokenCounts.ContainsKey("廠商名稱") &&
                      secrecy.ExpectedTokenCounts["廠商名稱"] == 3,
                    "manifest 保密同意書 廠商名稱x3（含黑字處）",
                    secrecy != null ? string.Join(",", secrecy.ExpectedTokenCounts.Select(
                        kv => kv.Key + "=" + kv.Value)) : "null");
                Check(secrecy != null && secrecy.Tokens.Contains("廠商名稱") &&
                      secrecy.Tokens.Contains("案件名稱") && secrecy.Tokens.Contains("機關名稱"),
                    "範本 token 直接掃描（不靠 tokens.txt）",
                    secrecy != null ? string.Join(",", secrecy.Tokens.OrderBy(x => x)) : "null");
                Check(store.Templates.All(t => t.LoadError == null), "範本皆可解析", null);
            }
            catch (Exception ex)
            {
                _failures++;
                Log.Add("FAIL 範本庫載入：" + ex.Message);
            }
            return store;
        }

        // ============================== 真實 xlsx ==============================
        static void TestRealXlsx(string baseDir, TemplateStore store)
        {
            string path = Path.Combine(baseDir, "標案資料.xlsx");
            if (!File.Exists(path)) { Log.Add("SKIP 真實 標案資料.xlsx 不存在"); return; }
            try
            {
                Dictionary<string, List<XlsxRow>> book = XlsxReader.Load(path);
                Check(book.ContainsKey("標案清單") && book.ContainsKey("公司資料"),
                    "真實 xlsx 工作表齊全", string.Join(",", book.Keys));
                List<XlsxRow> rows = book["標案清單"];
                List<string> headers = rows[0].Cells.Select(c => Util.Nfkc(c)).ToList();
                Check(headers.Contains("標案名稱") && headers.Contains("簽署日期"),
                    "真實 xlsx 表頭", string.Join("|", headers));
                if (store != null)
                {
                    foreach (TemplateInfo tpl in store.Templates)
                        Check(headers.Contains("產生:" + Util.Nfkc(tpl.BaseName)),
                            "真實 xlsx 有「產生:" + tpl.BaseName + "」欄", null);
                }
            }
            catch (Exception ex)
            {
                _failures++;
                Log.Add("FAIL 真實 xlsx 解析：" + ex.Message);
            }
        }

        // ============================== 端到端 ==============================
        static void TestEndToEnd(TemplateStore store, string tmpDir)
        {
            string longName = "116年度*資訊系統/維護:案<很長>" + new string('長', 60);
            string[][] tenders = new string[][]
            {
                new[] { "標案名稱", "機關名稱", "案件名稱", "簽署日期",
                        "產生:由所屬公司享有著作財產權與著作人格權同意書", "產生:保密切結書", "產生:保密同意書",
                        "備註", "廠商名稱", "負責人姓名", "聯絡電話", "廠商地址" },
                new[] { "測試標案Ａ全球資訊網維護案", "花蓮縣政府", "", "115年3月15日",
                        "是", "是", "是", "", "", "", "", "" },
                new[] { longName, "南投縣政府", "有&特殊<字>的案名", "116/2/28",
                        "是", "否", "是", "備註列",
                        "測試&有限<公司>", "王>小明", "@SCI@", "南投市中山路1號\n3樓" },
                new[] { "測試標案Ａ全球資訊網維護案", "重複機關", "", "115/1/1",
                        "是", "是", "是", "", "", "", "", "" },
                new[] { "CON", "測試機關", "", "115/1/1", "是", "否", "否", "", "", "", "", "" },
                new[] { "全部否的標案", "測試機關", "", "115/1/1", "否", "否", "否", "", "", "", "", "" }
            };
            string[][] company = new string[][]
            {
                new[] { "參數名稱", "值" },
                new[] { "廠商名稱", "恩潔測試股份有限公司" },
                new[] { "負責人姓名", "林測試" },
                new[] { "聯絡電話", "(07)999-8888" },
                new[] { "廠商地址", "高雄市苓雅區測試路99號" },
                new[] { "替換後文字顏色", "黑色" }
            };
            string xlsxPath = Path.Combine(tmpDir, "e2e.xlsx");
            File.WriteAllBytes(xlsxPath, BuildXlsx(tenders, company));

            GenerationPlan plan = Planner.BuildPlan(xlsxPath, store);
            Check(plan.GlobalErrors.Count == 0, "E2E 無全域錯誤", string.Join("；", plan.GlobalErrors));
            Check(plan.Rows.Count == 5, "E2E 5 個資料列", plan.Rows.Count.ToString());
            Check(!plan.KeepRed, "E2E 顏色設定=黑色", null);

            string outRoot = Path.Combine(tmpDir, "輸出");
            List<RowResult> results = plan.Rows.Select(r => Generator.GenerateRow(r, outRoot, false, plan.KeepRed)).ToList();

            // --- 列 1：正常列，3 個檔案 ---
            RowResult r1 = results[0];
            Check(r1.Errors.Count == 0 && r1.Generated.Count == 3, "列1 產生 3 檔",
                string.Join("；", r1.Errors) + "|" + r1.Generated.Count);
            string folder1 = Path.Combine(outRoot, "測試標案Ａ全球資訊網維護案");
            Check(Directory.Exists(folder1), "列1 資料夾", folder1);
            string body1 = ReadOdtBodyText(Path.Combine(folder1, "保密同意書.odt"));
            Check(body1.Contains("恩潔測試股份有限公司") && !body1.Contains("哈瑪星"),
                "列1 公司名替換、無殘留", null);
            Check(body1.Contains("測試標案Ａ全球資訊網維護案"), "列1 案件名稱 fallback=標案名稱", null);
            Check(body1.Contains("115") && body1.Contains("15"), "列1 簽署日期 115/3/15", null);
            Check(!ReadOdtContentXml(Path.Combine(folder1, "保密同意書.odt")).Contains("#FF0000"),
                "列1 黑色模式無紅字", null);

            // --- 列 2：非法字元＋截長＋跳脫＋換行＋科學記號電話，僅 2 檔 ---
            RowResult r2 = results[1];
            Check(r2.Errors.Count == 0 && r2.Generated.Count == 2, "列2 產生 2 檔（切結書=否）",
                string.Join("；", r2.Errors) + "|" + string.Join(",", r2.Generated));
            string folder2Name = r2.Plan.FolderName;
            Check(folder2Name.Length <= 60 && folder2Name.Contains("＊") && folder2Name.Contains("／") &&
                  folder2Name.Contains("：") && folder2Name.Contains("＜"),
                "列2 資料夾名清理與截長", folder2Name);
            string odt2 = Path.Combine(outRoot, folder2Name, "保密同意書.odt");
            string body2 = ReadOdtBodyText(odt2);
            string content2 = ReadOdtContentXml(odt2);
            Check(body2.Contains("有&特殊<字>的案名"), "列2 案名 &<> 正確跳脫還原", null);
            Check(body2.Contains("測試&有限<公司>") && body2.Contains("王>小明"), "列2 公司覆寫值與跳脫", null);
            Check(content2.Contains("&amp;") && content2.Contains("&lt;"), "列2 XML 實體跳脫存在", null);
            Check(content2.Contains("<text:line-break"), "列2 地址換行→line-break", null);
            Check(body2.Contains("912345678"), "列2 科學記號電話展開", null);

            // --- 列 3：重複標案名 → 錯誤 ---
            Check(results[2].Errors.Any(e => e.Contains("重複")), "列3 重複偵測",
                string.Join("；", results[2].Errors));

            // --- 列 4：保留裝置名 CON ---
            Check(results[3].Errors.Count == 0 && Directory.Exists(Path.Combine(outRoot, "＿CON")),
                "列4 CON→＿CON", string.Join("；", results[3].Errors));

            // --- 列 5：全部否 → 0 檔 ---
            Check(results[4].Generated.Count == 0 && results[4].Errors.Count == 0,
                "列5 全否 0 檔", null);

            // --- 第二輪：不覆寫 → 略過；覆寫 → 重產 ---
            RowResult again = Generator.GenerateRow(plan.Rows[0], outRoot, false, plan.KeepRed);
            Check(again.Skipped.Count == 3 && again.Generated.Count == 0, "重跑 略過已存在",
                again.Skipped.Count + "/" + again.Generated.Count);
            RowResult forced = Generator.GenerateRow(plan.Rows[0], outRoot, true, plan.KeepRed);
            Check(forced.Generated.Count == 3, "重跑 覆寫模式重產", forced.Generated.Count.ToString());
        }

        // ============================== 紅色模式 ==============================
        static void TestKeepRed(TemplateStore store, string tmpDir)
        {
            string[][] tenders = new string[][]
            {
                new[] { "標案名稱", "機關名稱", "簽署日期" },
                new[] { "紅色測試標案", "測試機關", "115/1/1" }
            };
            string[][] company = new string[][]
            {
                new[] { "廠商名稱", "紅色測試公司" },
                new[] { "負責人姓名", "紅" },
                new[] { "聯絡電話", "123" },
                new[] { "廠商地址", "紅色路1號" },
                new[] { "替換後文字顏色", "紅色" }
            };
            string xlsxPath = Path.Combine(tmpDir, "red.xlsx");
            File.WriteAllBytes(xlsxPath, BuildXlsx(tenders, company));
            GenerationPlan plan = Planner.BuildPlan(xlsxPath, store);
            Check(plan.KeepRed, "紅色設定解析", null);
            string outRoot = Path.Combine(tmpDir, "輸出紅");
            RowResult r = Generator.GenerateRow(plan.Rows[0], outRoot, false, plan.KeepRed);
            Check(r.Errors.Count == 0 && r.Generated.Count == 3, "紅色模式產生",
                string.Join("；", r.Errors));
            string content = ReadOdtContentXml(Path.Combine(outRoot, "紅色測試標案", "保密切結書.odt"));
            Check(content.Contains("#FF0000"), "紅色模式保留紅字", null);
        }

        // ============================== 缺參數 ==============================
        static void TestMissingParam(TemplateStore store, string tmpDir)
        {
            string[][] tenders = new string[][]
            {
                new[] { "標案名稱", "機關名稱", "簽署日期" },
                new[] { "缺參數標案", "測試機關", "115/1/1" }
            };
            string[][] company = new string[][]
            {
                new[] { "廠商名稱", "某公司" }
                // 缺 負責人姓名 / 聯絡電話 / 廠商地址
            };
            string xlsxPath = Path.Combine(tmpDir, "missing.xlsx");
            File.WriteAllBytes(xlsxPath, BuildXlsx(tenders, company));
            GenerationPlan plan = Planner.BuildPlan(xlsxPath, store);
            RowPlan rp = plan.Rows[0];
            Check(rp.Errors.Any(e => e.Contains("聯絡電話")) && rp.Errors.Any(e => e.Contains("負責人姓名")),
                "缺參數列出明確錯誤", string.Join("；", rp.Errors));
            RowResult r = Generator.GenerateRow(rp, Path.Combine(tmpDir, "輸出缺"), false, false);
            Check(r.Generated.Count == 0 && r.Errors.Count > 0, "缺參數不產生檔案", null);
        }

        // ============================== 範本竄改偵測 ==============================
        static void TestTamperDetection(TemplateStore store, string tmpDir)
        {
            string tamperDir = Path.Combine(tmpDir, "範本竄改");
            Directory.CreateDirectory(tamperDir);
            TemplateInfo src = store.Templates.First(t => t.BaseName == "保密切結書");

            List<KeyValuePair<string, byte[]>> entries = OdtWriter.ReadAllEntries(src.FullPath);
            for (int i = 0; i < entries.Count; i++)
            {
                if (entries[i].Key != "content.xml") continue;
                string xml = Encoding.UTF8.GetString(entries[i].Value);
                int pos = xml.IndexOf("${廠商名稱}", StringComparison.Ordinal);
                xml = xml.Substring(0, pos) + "被改掉了" + xml.Substring(pos + "${廠商名稱}".Length);
                entries[i] = new KeyValuePair<string, byte[]>("content.xml", Encoding.UTF8.GetBytes(xml));
            }
            File.WriteAllBytes(Path.Combine(tamperDir, src.FileName), OdtWriter.Repack(entries));
            File.Copy(Path.Combine(Path.GetDirectoryName(src.FullPath), "tokens.txt"),
                      Path.Combine(tamperDir, "tokens.txt"));

            TemplateStore tampered = TemplateStore.Load(tamperDir);
            Dictionary<string, string> values = new Dictionary<string, string>();
            foreach (string tk in tampered.Templates[0].ExpectedTokenCounts.Keys) values[tk] = "x";
            bool threw = false; string msg = "";
            try
            {
                HashSet<string> used;
                OdtWriter.Generate(tampered.Templates[0], values, false, out used);
            }
            catch (InvalidDataException ex) { threw = true; msg = ex.Message; }
            Check(threw && msg.Contains("編輯"), "範本竄改偵測", msg);
        }

        // ============================== tokens.txt 遺失仍可運作 ==============================
        static void TestTokensTxtMissing(string baseDir, string tmpDir)
        {
            // 複製範本但「不」複製 tokens.txt，模擬部署時漏掉或被刪除
            string dir = Path.Combine(tmpDir, "範本無manifest");
            Directory.CreateDirectory(dir);
            foreach (string odt in Directory.GetFiles(Path.Combine(baseDir, "範本"), "*.odt"))
                File.Copy(odt, Path.Combine(dir, Path.GetFileName(odt)));

            TemplateStore store = TemplateStore.Load(dir);
            TemplateInfo t = store.Templates.First(x => x.BaseName == "保密同意書");
            Check(t.ExpectedTokenCounts.Count == 0 && t.Tokens.Count > 0,
                "無 tokens.txt：manifest 為空但仍掃出範本 token",
                "expected=" + t.ExpectedTokenCounts.Count + " tokens=" + t.Tokens.Count);

            string[][] tenders = new string[][]
            {
                new[] { "標案名稱", "機關名稱", "簽署日期" },
                new[] { "無manifest測試", "測試機關", "115/1/1" }
            };
            string[][] company = new string[][]
            {
                new[] { "廠商名稱", "測試公司" }, new[] { "負責人姓名", "測試" },
                new[] { "聯絡電話", "12345" }, new[] { "廠商地址", "測試路1號" }
            };
            string xlsxPath = Path.Combine(tmpDir, "nomani.xlsx");
            File.WriteAllBytes(xlsxPath, BuildXlsx(tenders, company));
            GenerationPlan plan = Planner.BuildPlan(xlsxPath, store);
            RowResult r = Generator.GenerateRow(plan.Rows[0], Path.Combine(tmpDir, "輸出無mani"), false, false);
            Check(r.Errors.Count == 0 && r.Generated.Count == 3,
                "無 tokens.txt：仍正確產生 3 檔（非誤報缺參數）", string.Join("；", r.Errors));
        }

        // ============================== 值含 ${ 字面不被誤判 ==============================
        static void TestLiteralDollarBrace(TemplateStore store, string tmpDir)
        {
            string[][] tenders = new string[][]
            {
                new[] { "標案名稱", "機關名稱", "案件名稱", "簽署日期",
                        "產生:保密同意書" },
                new[] { "錢字號測試案", "測試機關", "AI${x}平台維護${案}", "115/1/1", "是" }
            };
            string[][] company = new string[][]
            {
                new[] { "廠商名稱", "測試公司" }, new[] { "負責人姓名", "測試" },
                new[] { "聯絡電話", "12345" }, new[] { "廠商地址", "測試路1號" }
            };
            string xlsxPath = Path.Combine(tmpDir, "dollar.xlsx");
            File.WriteAllBytes(xlsxPath, BuildXlsx(tenders, company));
            GenerationPlan plan = Planner.BuildPlan(xlsxPath, store);
            RowResult r = Generator.GenerateRow(plan.Rows[0], Path.Combine(tmpDir, "輸出錢"), false, false);
            Check(r.Errors.Count == 0 && r.Generated.Contains("保密同意書.odt"),
                "值含 ${ 字面不被誤判為未替換殘留", string.Join("；", r.Errors));
            string body = ReadOdtBodyText(Path.Combine(tmpDir, "輸出錢", "錢字號測試案", "保密同意書.odt"));
            Check(body.Contains("AI${x}平台維護${案}"), "值含 ${ 字面正確保留於文件", null);
        }

        // ============================== 注音導引（rPh）不污染值 ==============================
        static void TestPhoneticRuby()
        {
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            // <si><r><t>花蓮縣政府</t></r><rPh sb=0 eb=5><t>ㄅㄆㄇ…</t></rPh></si>
            XElement si = new XElement(ns + "si",
                new XElement(ns + "r", new XElement(ns + "t", "花蓮縣政府")),
                new XElement(ns + "rPh",
                    new XAttribute("sb", "0"), new XAttribute("eb", "5"),
                    new XElement(ns + "t", "ㄏㄨㄚㄌㄧㄢㄒㄧㄢㄓㄥㄈㄨ")));
            string tmp = Path.Combine(Path.GetTempPath(), "ph_test_ruby.xlsx");
            File.WriteAllBytes(tmp, BuildXlsxWithSharedString(si));
            try
            {
                Dictionary<string, List<XlsxRow>> book = XlsxReader.Load(tmp);
                string val = book["工作表1"][0].Cell(0);
                Check(val == "花蓮縣政府", "rPh 注音不污染 sharedString 儲存格值", "'" + val + "'");
            }
            finally { try { File.Delete(tmp); } catch { } }
        }

        // ============================== TemplateBuilder：新增範本正規化 ==============================
        static void TestTemplateBuilder(string tmpDir)
        {
            string odtPath = Path.Combine(tmpDir, "raw_template.odt");
            File.WriteAllBytes(odtPath, BuildColoredOdt());

            TemplateBuilder tb = TemplateBuilder.Load(odtPath);
            List<ColoredRun> runs = tb.ExtractRuns();
            Check(runs.Count == 3, "Builder 擷取 3 個彩色 run", runs.Count.ToString());
            ColoredRun addr = runs.FirstOrDefault(r => r.Text.Contains("中正路"));
            Check(addr != null && addr.Text == "台北市中正路1號", "Builder 相鄰紅 span 合併成地址",
                addr != null ? addr.Text : "null");
            Check(runs.Count(r => r.Category == "red") == 2 && runs.Count(r => r.Category == "blue") == 1,
                "Builder 紅/藍分類", string.Join(",", runs.Select(r => r.Category)));

            // 對應：公司名→廠商名稱、地址→廠商地址、機關→機關名稱
            Dictionary<string, string> map = new Dictionary<string, string>();
            foreach (ColoredRun r in runs)
            {
                if (r.Text.Contains("測試公司")) map[r.Id] = "廠商名稱";
                else if (r.Text.Contains("中正路")) map[r.Id] = "廠商地址";
                else if (r.Text.Contains("縣政府")) map[r.Id] = "機關名稱";
            }
            byte[] built = tb.Build(map);
            Check(TemplateBuilder.ValidateTemplateBytes(built).Count == 0, "Builder 產出 zip/xml 有效", null);

            // 存成範本 + 重寫 tokens.txt + 用 Generator 端到端產生
            string tdir = Path.Combine(tmpDir, "builder範本");
            Directory.CreateDirectory(tdir);
            File.WriteAllBytes(Path.Combine(tdir, "新範本.odt"), built);
            TemplateStore.RewriteTokensFile(tdir);
            TemplateStore store = TemplateStore.Load(tdir);
            TemplateInfo tpl = store.Templates[0];
            Check(tpl.Tokens.SetEquals(new HashSet<string> { "廠商名稱", "廠商地址", "機關名稱" }),
                "Builder 產出 token 集合正確", string.Join(",", tpl.Tokens.OrderBy(x => x)));

            Dictionary<string, string> values = new Dictionary<string, string>
            {
                { "廠商名稱", "端到端公司" }, { "廠商地址", "高雄市測試路9號" }, { "機關名稱", "測試縣政府" }
            };
            RowPlan rp = new RowPlan();
            rp.RowNumber = 2; rp.TenderName = "Builder端到端案"; rp.FolderName = "Builder端到端案";
            rp.Docs.Add(tpl);
            foreach (KeyValuePair<string, string> kv in values) rp.Values[kv.Key] = kv.Value;
            RowResult res = Generator.GenerateRow(rp, Path.Combine(tmpDir, "builder輸出"), false, false);
            Check(res.Errors.Count == 0 && res.Generated.Count == 1, "Builder 範本可被 Generator 產生",
                string.Join("；", res.Errors));
            string body = ReadOdtBodyText(Path.Combine(tmpDir, "builder輸出", "Builder端到端案", "新範本.odt"));
            Check(body.Contains("端到端公司") && body.Contains("高雄市測試路9號") && body.Contains("測試縣政府"),
                "Builder 範本替換值正確", null);
            Check(!body.Contains("測試公司股份有限公司"), "Builder 範本無殘留示範文字", null);
        }

        // ============================== DocxToOdt：Word 轉範本來源 ==============================
        static void TestDocxToOdt(string tmpDir)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            Func<string, string, bool, XElement> crun = delegate(string color, string txt, bool bold)
            {
                XElement rPr = new XElement(w + "rPr");
                if (color != null) rPr.Add(new XElement(w + "color", new XAttribute(w + "val", color)));
                if (bold) rPr.Add(new XElement(w + "b"));
                XElement r = new XElement(w + "r");
                if (rPr.HasElements) r.Add(rPr);
                r.Add(new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), txt));
                return r;
            };
            Func<string, XElement> listPara = delegate(string txt)
            {
                return new XElement(w + "p",
                    new XElement(w + "pPr", new XElement(w + "numPr",
                        new XElement(w + "ilvl", new XAttribute(w + "val", "0")),
                        new XElement(w + "numId", new XAttribute(w + "val", "1")))),
                    new XElement(w + "r", new XElement(w + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"), txt)));
            };

            XElement body = new XElement(w + "body",
                new XElement(w + "p",
                    crun("FF0000", "測試公司股份有限公司", false),
                    crun(null, "，委派至", false),
                    crun("0070C0", "某某縣政府", true),
                    crun(null, "辦理。", false)),
                listPara("第一項"),
                listPara("第二項"),
                new XElement(w + "p",
                    new XElement(w + "pPr", new XElement(w + "jc", new XAttribute(w + "val", "center"))),
                    crun(null, "報告完畢", false)),
                // 區塊層級 sdt（內容控制項）包住的段落——不可被漏掉
                new XElement(w + "sdt", new XElement(w + "sdtContent",
                    new XElement(w + "p", crun(null, "內容控制項段落", false)))),
                // 超連結內含追蹤修訂插入(ins)的 run——巢狀容器不可漏字
                new XElement(w + "p",
                    new XElement(w + "hyperlink",
                        new XElement(w + "ins", new XAttribute(w + "id", "1"),
                            crun(null, "連結插入文字", false)))),
                // 簡單功能變數（w:fldSimple）的結果 run——不可漏字
                new XElement(w + "p",
                    new XElement(w + "fldSimple", new XAttribute(w + "instr", " REF x "),
                        crun(null, "功能變數文字", false))),
                // 由段落樣式（w:pStyle）帶出編號的清單段落
                new XElement(w + "p",
                    new XElement(w + "pPr", new XElement(w + "pStyle", new XAttribute(w + "val", "ListNumber"))),
                    crun(null, "樣式編號項", false)),
                new XElement(w + "sectPr",
                    new XElement(w + "pgSz", new XAttribute(w + "w", "11906"), new XAttribute(w + "h", "16838")),
                    new XElement(w + "pgMar", new XAttribute(w + "top", "1440"), new XAttribute(w + "right", "1440"),
                        new XAttribute(w + "bottom", "1440"), new XAttribute(w + "left", "1440"))));
            XDocument document = new XDocument(new XElement(w + "document", body));
            XDocument numbering = new XDocument(new XElement(w + "numbering",
                new XElement(w + "abstractNum", new XAttribute(w + "abstractNumId", "0"),
                    new XElement(w + "lvl", new XAttribute(w + "ilvl", "0"),
                        new XElement(w + "numFmt", new XAttribute(w + "val", "decimal")),
                        new XElement(w + "lvlText", new XAttribute(w + "val", "%1.")))),
                new XElement(w + "num", new XAttribute(w + "numId", "1"),
                    new XElement(w + "abstractNumId", new XAttribute(w + "val", "0")))));
            XDocument styles = new XDocument(new XElement(w + "styles",
                new XElement(w + "style", new XAttribute(w + "type", "paragraph"),
                    new XAttribute(w + "styleId", "ListNumber"),
                    new XElement(w + "pPr", new XElement(w + "numPr",
                        new XElement(w + "ilvl", new XAttribute(w + "val", "0")),
                        new XElement(w + "numId", new XAttribute(w + "val", "1")))))));

            byte[] docx = BuildDocx(document, numbering, styles);
            List<string> warnings;
            byte[] odt = DocxToOdt.ConvertBytes(docx, out warnings);
            Check(TemplateBuilder.ValidateTemplateBytes(odt).Count == 0, "Docx→Odt 產出 zip/xml 有效",
                string.Join("；", TemplateBuilder.ValidateTemplateBytes(odt)));

            string raw = Path.Combine(tmpDir, "docx_raw.odt");
            File.WriteAllBytes(raw, odt);

            // styles.xml 存在、well-formed、A4 頁寬
            string stylesXml = ReadOdtEntry(raw, "styles.xml");
            Check(stylesXml != null && stylesXml.Contains("master-page") && stylesXml.Contains("page-width=\"21"),
                "Docx→Odt styles.xml 有頁面版面（A4）", stylesXml == null ? "null" :
                (stylesXml.Contains("page-width=\"21") ? "ok" : "無 21cm 頁寬"));

            // 用既有 TemplateBuilder 擷取彩色 run（紅/藍兩段）
            TemplateBuilder tb = TemplateBuilder.Load(raw);
            List<ColoredRun> runs = tb.ExtractRuns();
            Check(runs.Count == 2, "Docx→Odt 擷取 2 個彩色 run",
                runs.Count + "：" + string.Join(",", runs.Select(r => r.Text)));
            Check(runs.Any(r => r.Text == "測試公司股份有限公司" && r.Category == "red")
                && runs.Any(r => r.Text == "某某縣政府" && r.Category == "blue"),
                "Docx→Odt 紅/藍 run 文字與分類正確",
                string.Join("|", runs.Select(r => r.Text + "=" + r.Category)));

            // 正規化 → 存範本 → 端到端產生
            Dictionary<string, string> map = new Dictionary<string, string>();
            foreach (ColoredRun r in runs)
            {
                if (r.Text.Contains("測試公司")) map[r.Id] = "廠商名稱";
                else if (r.Text.Contains("縣政府")) map[r.Id] = "機關名稱";
            }
            byte[] built = tb.Build(map);
            Check(TemplateBuilder.ValidateTemplateBytes(built).Count == 0, "Docx→Odt 正規化後仍有效", null);

            string tdir = Path.Combine(tmpDir, "docx範本");
            Directory.CreateDirectory(tdir);
            File.WriteAllBytes(Path.Combine(tdir, "Word範本.odt"), built);
            TemplateStore.RewriteTokensFile(tdir);
            TemplateStore store = TemplateStore.Load(tdir);
            TemplateInfo tpl = store.Templates[0];
            Check(tpl.Tokens.SetEquals(new HashSet<string> { "廠商名稱", "機關名稱" }),
                "Docx→Odt 範本 token 集合正確", string.Join(",", tpl.Tokens.OrderBy(x => x)));

            RowPlan rp = new RowPlan();
            rp.RowNumber = 2; rp.TenderName = "Word端到端案"; rp.FolderName = "Word端到端案"; rp.Docs.Add(tpl);
            rp.Values["廠商名稱"] = "轉入公司"; rp.Values["機關名稱"] = "轉入縣政府";
            RowResult res = Generator.GenerateRow(rp, Path.Combine(tmpDir, "docx輸出"), false, false);
            Check(res.Errors.Count == 0 && res.Generated.Count == 1, "Docx→Odt 範本可被 Generator 產生",
                string.Join("；", res.Errors));
            string outBody = ReadOdtBodyText(Path.Combine(tmpDir, "docx輸出", "Word端到端案", "Word範本.odt"));
            Check(outBody.Contains("轉入公司") && outBody.Contains("轉入縣政府"),
                "Docx→Odt 替換值正確", null);
            Check(outBody.Contains("第一項") && outBody.Contains("第二項") && outBody.Contains("報告完畢"),
                "Docx→Odt 清單與段落文字保留", null);
            string outContent = ReadOdtContentXml(Path.Combine(tmpDir, "docx輸出", "Word端到端案", "Word範本.odt"));
            Check(outContent.Contains("<text:list"), "Docx→Odt 清單轉為 text:list", null);
            Check(outContent.Contains("text-align=\"center\""), "Docx→Odt 置中對齊保留", null);
            Check(outContent.Contains("style:num-suffix=\".\"") && !outContent.Contains("text:num-suffix"),
                "Docx→Odt 編號後綴用 style:num-suffix（ODF 合規，非 text:）",
                outContent.Contains("text:num-suffix") ? "仍有 text:num-suffix" : "ok");
            Check(outBody.Contains("內容控制項段落"), "Docx→Odt 區塊 sdt 段落不遺失", null);
            Check(outBody.Contains("連結插入文字"), "Docx→Odt 超連結內巢狀 run 不遺失", null);
            Check(outBody.Contains("功能變數文字"), "Docx→Odt fldSimple 結果 run 不遺失", null);
            Check(outBody.Contains("樣式編號項"), "Docx→Odt pStyle 清單段落文字保留", null);
            int listCount = outContent.Split(new string[] { "<text:list " }, StringSplitOptions.None).Length - 1;
            Check(listCount >= 2, "Docx→Odt pStyle 段落亦成為 text:list（≥2 個清單）", "listCount=" + listCount);

            // 不支援結構（表格）→ 明確報錯
            XDocument tableDoc = new XDocument(new XElement(w + "document",
                new XElement(w + "body",
                    new XElement(w + "tbl", new XElement(w + "tr", new XElement(w + "tc",
                        new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", "x")))))))));
            byte[] tdocx = BuildDocx(tableDoc, null);
            bool threw = false; string msg = "";
            try { List<string> ww; DocxToOdt.ConvertBytes(tdocx, out ww); }
            catch (InvalidDataException ex) { threw = true; msg = ex.Message; }
            Check(threw && msg.Contains("表格"), "Docx→Odt 含表格→明確報錯", msg);
        }

        // ============================== 合併時穿透的空白不殘留於 ${token} 旁 ==============================
        static void TestMergeWhitespaceFix(string tmpDir)
        {
            XNamespace o = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
            XNamespace t = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
            XNamespace s = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
            XNamespace fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
            // 兩段紅色 span，中間只隔一個「未上色的空白文字節點」，後接一般文字
            XDocument doc = new XDocument(new XDeclaration("1.0", "UTF-8", null),
                new XElement(o + "document-content",
                    new XElement(o + "automatic-styles",
                        new XElement(s + "style", new XAttribute(s + "name", "C_RED"), new XAttribute(s + "family", "text"),
                            new XElement(s + "text-properties", new XAttribute(fo + "color", "#FF0000")))),
                    new XElement(o + "body", new XElement(o + "text",
                        new XElement(t + "p",
                            new XElement(t + "span", new XAttribute(t + "style-name", "C_RED"), "甲"),
                            " ",
                            new XElement(t + "span", new XAttribute(t + "style-name", "C_RED"), "乙"),
                            "後續")))));
            byte[] content;
            using (MemoryStream ms = new MemoryStream())
            {
                XmlWriterSettings set = new XmlWriterSettings();
                set.Encoding = new UTF8Encoding(false); set.Indent = false;
                using (XmlWriter xw = XmlWriter.Create(ms, set)) { doc.Save(xw); }
                content = ms.ToArray();
            }
            List<KeyValuePair<string, byte[]>> entries = new List<KeyValuePair<string, byte[]>>();
            entries.Add(new KeyValuePair<string, byte[]>("mimetype",
                Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.text")));
            entries.Add(new KeyValuePair<string, byte[]>("content.xml", content));
            string raw = Path.Combine(tmpDir, "ws_raw.odt");
            File.WriteAllBytes(raw, OdtWriter.Repack(entries));

            TemplateBuilder tb = TemplateBuilder.Load(raw);
            List<ColoredRun> runs = tb.ExtractRuns();
            Check(runs.Count == 1 && runs[0].Text == "甲乙",
                "空白合併：兩段同色 span 穿透空白合成一段", runs.Count + ":" + (runs.Count > 0 ? runs[0].Text : ""));
            Dictionary<string, string> map = new Dictionary<string, string>();
            foreach (ColoredRun r in runs) map[r.Id] = "廠商名稱";
            byte[] built = tb.Build(map);

            string tdir = Path.Combine(tmpDir, "ws範本");
            Directory.CreateDirectory(tdir);
            File.WriteAllBytes(Path.Combine(tdir, "空白測試.odt"), built);
            TemplateStore.RewriteTokensFile(tdir);
            TemplateStore store = TemplateStore.Load(tdir);
            RowPlan rp = new RowPlan();
            rp.RowNumber = 2; rp.TenderName = "空白案"; rp.FolderName = "空白案"; rp.Docs.Add(store.Templates[0]);
            rp.Values["廠商名稱"] = "值";
            RowResult res = Generator.GenerateRow(rp, Path.Combine(tmpDir, "ws輸出"), false, false);
            Check(res.Errors.Count == 0 && res.Generated.Count == 1, "空白合併：可產生", string.Join("；", res.Errors));
            string body = ReadOdtBodyText(Path.Combine(tmpDir, "ws輸出", "空白案", "空白測試.odt"));
            Check(body.Contains("值後續") && !body.Contains("值 後續"),
                "空白合併：${token} 旁無殘留空格", "'" + body + "'");
        }

        /// <summary>把 document.xml（與可選 numbering.xml / styles.xml）打包成最小 .docx（僅含轉換器會讀的 part）。</summary>
        static byte[] BuildDocx(XDocument document, XDocument numbering, XDocument styles = null)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (ZipArchive zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
                {
                    WriteXml(zip, "word/document.xml", document);
                    if (numbering != null) WriteXml(zip, "word/numbering.xml", numbering);
                    if (styles != null) WriteXml(zip, "word/styles.xml", styles);
                }
                return ms.ToArray();
            }
        }

        static string ReadOdtEntry(string path, string entryName)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                ZipArchiveEntry e = zip.GetEntry(entryName);
                if (e == null) return null;
                using (StreamReader sr = new StreamReader(e.Open(), Encoding.UTF8)) return sr.ReadToEnd();
            }
        }

        // ============================== XlsxEditor：加欄/加列 ==============================
        static void TestXlsxEditor(string baseDir, string tmpDir)
        {
            string src = Path.Combine(baseDir, "標案資料.範例.xlsx");
            if (!File.Exists(src)) src = Path.Combine(baseDir, "標案資料.xlsx");
            if (!File.Exists(src)) { Log.Add("SKIP XlsxEditor：找不到範例 xlsx"); return; }

            string work = Path.Combine(tmpDir, "editor.xlsx");
            File.Copy(src, work, true);
            int before = XlsxReader.Load(work).First(kv => Util.Nfkc(kv.Key) == "標案清單")
                .Value.Skip(1).Count(r => r.Cells.Any(c => (c ?? "").Trim() != ""));

            XlsxAdditions add = new XlsxAdditions();
            add.TenderColumns.Add("履約期限");
            add.TenderColumns.Add("產生:測試新範本");
            add.DropdownColumns.Add("產生:測試新範本");
            add.CompanyParams.Add("統一編號");
            XlsxEditor.Apply(work, add);

            Dictionary<string, List<XlsxRow>> book = XlsxReader.Load(work);
            List<XlsxRow> tenders = book.First(kv => Util.Nfkc(kv.Key) == "標案清單").Value;
            List<XlsxRow> company = book.First(kv => Util.Nfkc(kv.Key) == "公司資料").Value;
            HashSet<string> headers = new HashSet<string>(tenders[0].Cells.Select(c => Util.Nfkc(c)));
            Check(headers.Contains("履約期限") && headers.Contains("產生:測試新範本"),
                "XlsxEditor 標案清單新增 2 欄", string.Join("|", headers));
            HashSet<string> keys = new HashSet<string>(company.Select(r => Util.Nfkc(r.Cell(0))));
            Check(keys.Contains("統一編號"), "XlsxEditor 公司資料新增列", string.Join("|", keys));
            int after = tenders.Skip(1).Count(r => r.Cells.Any(c => (c ?? "").Trim() != ""));
            Check(after == before, "XlsxEditor 既有標案資料列數不變", before + "→" + after);
            // 既有值仍在
            Check(tenders.Skip(1).Any(r => r.Cells.Any(c => (c ?? "").Contains("觀光資訊網"))),
                "XlsxEditor 既有資料內容保留", null);
        }

        // ============================== 重新產生：先清空資料夾 ==============================
        static void TestRegenerateClearsFolder(TemplateStore store, string tmpDir)
        {
            string outRoot = Path.Combine(tmpDir, "regen輸出");
            TemplateInfo t1 = store.Templates.First(t => t.BaseName == "保密切結書");
            TemplateInfo t2 = store.Templates.First(t => t.BaseName == "保密同意書");
            Dictionary<string, string> vals = new Dictionary<string, string>
            {
                { "廠商名稱", "甲公司" }, { "負責人姓名", "乙" }, { "聯絡電話", "12345" }, { "廠商地址", "丙路1號" },
                { "機關名稱", "丁機關" }, { "案件名稱", "戊案" }, { "簽署年", "115" }, { "簽署月", "1" }, { "簽署日", "1" }
            };
            string folder = Path.Combine(outRoot, "清空測試");

            // 先產生 2 份文件
            RowPlan rp2 = new RowPlan();
            rp2.RowNumber = 2; rp2.TenderName = "清空測試"; rp2.FolderName = "清空測試";
            rp2.Docs.Add(t1); rp2.Docs.Add(t2);
            foreach (KeyValuePair<string, string> kv in vals) rp2.Values[kv.Key] = kv.Value;
            Generator.GenerateRow(rp2, outRoot, false, false);
            Check(File.Exists(Path.Combine(folder, t1.FileName)) && File.Exists(Path.Combine(folder, t2.FileName)),
                "重產前置：先產生 2 檔", null);

            // 改成只選 1 份，用「重新產生（overwrite）」→ 另一份舊檔應被清掉
            RowPlan rp1 = new RowPlan();
            rp1.RowNumber = 2; rp1.TenderName = "清空測試"; rp1.FolderName = "清空測試";
            rp1.Docs.Add(t1);
            foreach (KeyValuePair<string, string> kv in vals) rp1.Values[kv.Key] = kv.Value;
            RowResult r = Generator.GenerateRow(rp1, outRoot, true, false);
            Check(r.Errors.Count == 0 && File.Exists(Path.Combine(folder, t1.FileName))
                && !File.Exists(Path.Combine(folder, t2.FileName)),
                "重新產生：清空後只剩目前選取的文件（舊檔已移除）", string.Join(";", r.Errors));
            Check(Directory.GetFiles(folder, "*.odt").Length == 1, "重新產生：資料夾內恰 1 個 odt",
                Directory.GetFiles(folder, "*.odt").Length.ToString());

            // 對照：不勾 overwrite 時不清空（既有檔保留、只補缺）
            RowResult r2 = Generator.GenerateRow(rp2, outRoot, false, false);
            Check(r2.Skipped.Contains(t1.FileName) && File.Exists(Path.Combine(folder, t2.FileName)),
                "非重新產生：不清空、既有略過只補缺", "skipped=" + string.Join(",", r2.Skipped));
        }

        // ============================== XlsxEditor：插入位置 + 移除 ==============================
        static void TestXlsxInsertPositionAndRemove(string baseDir, string tmpDir)
        {
            string src = Path.Combine(baseDir, "標案資料.範例.xlsx");
            if (!File.Exists(src)) src = Path.Combine(baseDir, "標案資料.xlsx");
            if (!File.Exists(src)) { Log.Add("SKIP 插入/移除：找不到範例 xlsx"); return; }

            string work = Path.Combine(tmpDir, "insrem.xlsx");
            File.Copy(src, work, true);
            List<string> orig = TenderHeaders(work);
            int before = TenderDataRowCount(work);

            // 插入：產生欄 + 一個每案參數欄（服務層順序：產生欄先）
            XlsxAdditions add = new XlsxAdditions();
            add.TenderColumns.Add("產生:新範本A");
            add.DropdownColumns.Add("產生:新範本A");
            add.TenderColumns.Add("履約地點");
            XlsxEditor.Apply(work, add);

            List<string> h = TenderHeaders(work);
            int iGenLast = h.FindIndex(x => x == "產生:保密同意書");
            int iNewGen = h.FindIndex(x => x == "產生:新範本A");
            int iLoc = h.FindIndex(x => x == "履約地點");
            int iNote = h.FindIndex(x => x == "備註");
            Check(iNewGen == iGenLast + 1, "插入：產生欄緊接最後一個產生欄（群組連續）",
                "產生:保密同意書@" + iGenLast + " 新產生@" + iNewGen);
            Check(iLoc == iNewGen + 1 && iNote == iLoc + 1, "插入：參數欄接在產生欄後、在備註之前",
                "新產生@" + iNewGen + " 履約@" + iLoc + " 備註@" + iNote);
            // ListObject 的 tableColumn 順序必須與表頭一致（否則 Excel 跳修復）
            List<string> tcols = TableColumnNames(work);
            Check(tcols.SequenceEqual(h), "插入：tableColumn 順序與表頭一致",
                "table=[" + string.Join(",", tcols) + "]");
            Check(TenderDataRowCount(work) == before, "插入：既有資料列數不變", null);
            Check(CompanyKeysOf(work).Contains("履約地點") == false, "插入：每案欄不誤入公司資料", null);

            // 連續再插一個 → 產生欄仍連續
            XlsxAdditions add2 = new XlsxAdditions();
            add2.TenderColumns.Add("產生:新範本B");
            add2.DropdownColumns.Add("產生:新範本B");
            XlsxEditor.Apply(work, add2);
            List<string> h2 = TenderHeaders(work);
            int a = h2.FindIndex(x => x == "產生:新範本A");
            int b = h2.FindIndex(x => x == "產生:新範本B");
            Check(b == a + 1, "連續插入：產生欄群組仍連續", "A@" + a + " B@" + b);

            // 移除「新範本A」的欄 + 一個每案參數欄
            XlsxRemovals rem = new XlsxRemovals();
            rem.TenderColumns.Add("產生:新範本A");
            rem.TenderColumns.Add("履約地點");
            XlsxEditor.Remove(work, rem);
            List<string> h3 = TenderHeaders(work);
            Check(!h3.Contains("產生:新範本A") && !h3.Contains("履約地點"), "移除：欄位已消失", string.Join("|", h3));
            Check(h3.Contains("產生:新範本B") && h3.Contains("備註") && h3.Contains("廠商地址"),
                "移除：其他欄位保留", null);
            Check(TenderDataRowCount(work) == before, "移除：既有資料列數不變", null);

            // 全部移回原狀
            XlsxRemovals rem2 = new XlsxRemovals();
            rem2.TenderColumns.Add("產生:新範本B");
            XlsxEditor.Remove(work, rem2);
            List<string> h4 = TenderHeaders(work);
            Check(h4.Count == orig.Count && h4.SequenceEqual(orig), "移除：回復與原始相同欄序",
                "orig=" + orig.Count + " now=" + h4.Count);
        }

        static void TestRemoveTemplateService(string baseDir, string tmpDir)
        {
            string src = Path.Combine(baseDir, "標案資料.範例.xlsx");
            if (!File.Exists(src)) src = Path.Combine(baseDir, "標案資料.xlsx");
            if (!File.Exists(src)) { Log.Add("SKIP 移除範本：找不到範例 xlsx"); return; }

            string svc = Path.Combine(tmpDir, "svcrem");
            Directory.CreateDirectory(Path.Combine(svc, "範本"));
            File.Copy(src, Path.Combine(svc, "標案資料.xlsx"), true);
            // 先放兩份真實範本＋新增一份含獨有參數的
            File.Copy(Path.Combine(baseDir, "範本", "保密切結書.odt"), Path.Combine(svc, "範本", "保密切結書.odt"));
            File.Copy(Path.Combine(baseDir, "範本", "保密同意書.odt"), Path.Combine(svc, "範本", "保密同意書.odt"));
            string odt = Path.Combine(tmpDir, "svcrem_raw.odt");
            File.WriteAllBytes(odt, BuildColoredOdt());
            TemplateBuilder tb = TemplateBuilder.Load(odt);
            Dictionary<string, string> map = new Dictionary<string, string>();
            foreach (ColoredRun r in tb.ExtractRuns())
            {
                if (r.Text.Contains("測試公司")) map[r.Id] = "統編A";     // 全新「公司固定」參數
                else if (r.Text.Contains("中正路")) map[r.Id] = "獨有地點"; // 全新「每案」參數
                else if (r.Text.Contains("縣政府")) map[r.Id] = "機關名稱"; // 共用（既有）
            }
            AddTemplateResult addRes = AddTemplateService.Commit(odt, "待移除範本", map,
                new Dictionary<string, bool> { { "統編A", true }, { "獨有地點", false } }, svc, false);
            Check(addRes.Ok, "移除前：先建立成功", addRes.Error);
            Check(CompanyKeysOf(Path.Combine(svc, "標案資料.xlsx")).Contains("統編A"),
                "移除前：公司固定參數已入公司資料", null);

            // 預覽移除：獨有每案欄→標案清單、公司固定參數→公司資料、共用不列
            RemovalPlan prev = AddTemplateService.PreviewRemoval("待移除範本", svc);
            Check(prev.Error == null && prev.TenderColumns.Contains("產生:待移除範本")
                && prev.TenderColumns.Contains("獨有地點")
                && prev.CompanyParams.Contains("統編A")
                && !prev.TenderColumns.Contains("機關名稱"),
                "移除預覽：獨有每案欄+公司參數正確分類（共用不列）",
                "tender=" + string.Join(",", prev.TenderColumns) + " company=" + string.Join(",", prev.CompanyParams));

            RemovalPlan rres = AddTemplateService.RemoveTemplate("待移除範本", svc);
            Check(rres.Ok, "移除範本成功", rres.Error);
            Check(!File.Exists(Path.Combine(svc, "範本", "待移除範本.odt")), "移除：範本檔已刪", null);
            List<string> hdr = TenderHeaders(Path.Combine(svc, "標案資料.xlsx"));
            Check(!hdr.Contains("產生:待移除範本") && !hdr.Contains("獨有地點"), "移除：標案清單欄已消失", null);
            Check(!CompanyKeysOf(Path.Combine(svc, "標案資料.xlsx")).Contains("統編A"),
                "移除：公司資料獨有參數列已消失（不再遺留）", null);
            Check(hdr.Contains("機關名稱"), "移除：共用欄保留", null);
        }

        // 移除「最後一個產生欄」後，是/否 驗證範圍要縮小、不可蔓延到備註欄
        static void TestValidationShrinkOnRemove(string baseDir, string tmpDir)
        {
            string src = Path.Combine(baseDir, "標案資料.範例.xlsx");
            if (!File.Exists(src)) src = Path.Combine(baseDir, "標案資料.xlsx");
            if (!File.Exists(src)) { Log.Add("SKIP 驗證縮範圍：找不到範例 xlsx"); return; }
            string work = Path.Combine(tmpDir, "valshrink.xlsx");
            File.Copy(src, work, true);

            XlsxRemovals rem = new XlsxRemovals();
            rem.TenderColumns.Add("產生:保密同意書");   // 最後一個產生欄（G）
            XlsxEditor.Remove(work, rem);

            List<string> sqrefs = TenderValidationSqrefs(work);
            string all = string.Join(" ", sqrefs);
            Check(all.Contains("E2:F4") && !all.Contains("G"),
                "移除末產生欄：是/否驗證縮為 E2:F4（未蔓延到備註）", "sqref=" + all);
        }

        static List<string> TenderHeaders(string xlsxPath)
        {
            return XlsxReader.Load(xlsxPath).First(kv => Util.Nfkc(kv.Key) == "標案清單")
                .Value[0].Cells.Select(c => Util.Nfkc(c)).ToList();
        }
        // 讀 xl/tables/table1.xml 的 tableColumn name（依序），驗證與表頭一致。
        static List<string> TableColumnNames(string xlsxPath)
        {
            using (FileStream fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                ZipArchiveEntry e = zip.Entries.FirstOrDefault(x => x.FullName.StartsWith("xl/tables/") && x.FullName.EndsWith(".xml"));
                if (e == null) return new List<string>();
                using (Stream s = e.Open())
                {
                    XDocument d = XDocument.Load(s);
                    return d.Descendants().Where(x => x.Name.LocalName == "tableColumn")
                        .Select(x => Util.Nfkc((string)x.Attribute("name"))).ToList();
                }
            }
        }
        static HashSet<string> CompanyKeysOf(string xlsxPath)
        {
            return new HashSet<string>(XlsxReader.Load(xlsxPath).First(kv => Util.Nfkc(kv.Key) == "公司資料")
                .Value.Select(r => Util.Nfkc(r.Cell(0))));
        }
        static int TenderDataRowCount(string xlsxPath)
        {
            return XlsxReader.Load(xlsxPath).First(kv => Util.Nfkc(kv.Key) == "標案清單")
                .Value.Skip(1).Count(r => r.Cells.Any(c => (c ?? "").Trim() != ""));
        }
        // 讀標案清單 worksheet 的 dataValidation sqref（找含 sheetData 且有 dataValidation 的工作表）
        static List<string> TenderValidationSqrefs(string xlsxPath)
        {
            using (FileStream fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry e in zip.Entries.Where(x => x.FullName.StartsWith("xl/worksheets/") && x.FullName.EndsWith(".xml")))
                {
                    using (Stream s = e.Open())
                    {
                        XDocument d = XDocument.Load(s);
                        List<string> sq = d.Descendants().Where(x => x.Name.LocalName == "dataValidation")
                            .Select(x => (string)x.Attribute("sqref")).Where(v => v != null).ToList();
                        if (sq.Count > 0) return sq;
                    }
                }
            }
            return new List<string>();
        }

        // ============================== AddTemplateService：完整流程（含 Excel 更新） ==============================
        static void TestAddTemplateService(string baseDir, string tmpDir)
        {
            string src = Path.Combine(baseDir, "標案資料.範例.xlsx");
            if (!File.Exists(src)) src = Path.Combine(baseDir, "標案資料.xlsx");
            if (!File.Exists(src)) { Log.Add("SKIP AddTemplateService：找不到範例 xlsx"); return; }

            // 建一個獨立的 baseDir：範本\（空）＋ 標案資料.xlsx
            string svc = Path.Combine(tmpDir, "svc");
            Directory.CreateDirectory(Path.Combine(svc, "範本"));
            File.Copy(src, Path.Combine(svc, "標案資料.xlsx"), true);
            string odt = Path.Combine(tmpDir, "svc_raw.odt");
            File.WriteAllBytes(odt, BuildColoredOdt());

            TemplateBuilder tb = TemplateBuilder.Load(odt);
            List<ColoredRun> runs = tb.ExtractRuns();
            Dictionary<string, string> map = new Dictionary<string, string>();
            foreach (ColoredRun r in runs)
            {
                if (r.Text.Contains("測試公司")) map[r.Id] = "統一編號";   // 全新公司固定參數
                else if (r.Text.Contains("中正路")) map[r.Id] = "履約地點"; // 全新每案參數
                else if (r.Text.Contains("縣政府")) map[r.Id] = "機關名稱"; // 沿用既有
            }
            Dictionary<string, bool> newType = new Dictionary<string, bool>
            { { "統一編號", true }, { "履約地點", false } };

            AddTemplateResult res = AddTemplateService.Commit(odt, "服務測試範本", map, newType, svc, false);
            Check(res.Ok, "Service 建立成功", res.Error);
            if (!res.Ok) return;

            // 覆寫同名範本：不可產生重複欄（sharedString 表頭去重必須有效）
            AddTemplateResult res2 = AddTemplateService.Commit(odt, "服務測試範本", map, newType, svc, true);
            Check(res2.Ok, "Service 覆寫同名成功", res2.Error);
            List<string> hdr2 = XlsxReader.Load(Path.Combine(svc, "標案資料.xlsx"))
                .First(kv => Util.Nfkc(kv.Key) == "標案清單").Value[0].Cells.Select(c => Util.Nfkc(c)).ToList();
            Check(hdr2.Count(h => h == "產生:服務測試範本") == 1, "Service 覆寫不重複加「產生」欄",
                hdr2.Count(h => h == "產生:服務測試範本").ToString());
            Check(hdr2.Count(h => h == "履約地點") == 1, "Service 覆寫不重複加參數欄",
                hdr2.Count(h => h == "履約地點").ToString());

            // 參數名稱含 } 應被拒（否則範本永遠無法產生）
            Dictionary<string, string> badMap = new Dictionary<string, string>();
            foreach (ColoredRun r in runs) { badMap[r.Id] = "壞}名"; break; }
            AddTemplateResult resBad = AddTemplateService.Commit(odt, "壞參數範本", badMap, new Dictionary<string, bool>(), svc, false);
            Check(!resBad.Ok && resBad.Error.Contains("不允許"), "Service 拒絕含 } 的參數名", resBad.Error);

            Check(File.Exists(Path.Combine(svc, "範本", "服務測試範本.odt")), "Service 範本已存檔", null);

            // Excel：新每案欄、產生欄、公司列
            Dictionary<string, List<XlsxRow>> book = XlsxReader.Load(Path.Combine(svc, "標案資料.xlsx"));
            HashSet<string> headers = new HashSet<string>(
                book.First(kv => Util.Nfkc(kv.Key) == "標案清單").Value[0].Cells.Select(c => Util.Nfkc(c)));
            Check(headers.Contains("履約地點"), "Service 加了每案欄 履約地點", null);
            Check(headers.Contains("產生:服務測試範本"), "Service 加了產生欄", null);
            Check(!headers.Contains("機關名稱") == false, "Service 沿用既有 機關名稱 欄", null);
            HashSet<string> compKeys = new HashSet<string>(
                book.First(kv => Util.Nfkc(kv.Key) == "公司資料").Value.Select(r => Util.Nfkc(r.Cell(0))));
            Check(compKeys.Contains("統一編號"), "Service 加了公司參數 統一編號", null);

            // 端到端：Generator 能用新範本產生（提供三個參數值）
            TemplateStore store = TemplateStore.Load(Path.Combine(svc, "範本"));
            TemplateInfo tpl = store.Templates.First(t => t.BaseName == "服務測試範本");
            RowPlan rp = new RowPlan();
            rp.RowNumber = 2; rp.TenderName = "Svc端到端"; rp.FolderName = "Svc端到端"; rp.Docs.Add(tpl);
            rp.Values["統一編號"] = "12345678";
            rp.Values["履約地點"] = "台中市";
            rp.Values["機關名稱"] = "端到端縣政府";
            RowResult gr = Generator.GenerateRow(rp, Path.Combine(tmpDir, "svc輸出"), false, false);
            Check(gr.Errors.Count == 0 && gr.Generated.Count == 1, "Service 新範本可產生文件",
                string.Join("；", gr.Errors));
            string body = ReadOdtBodyText(Path.Combine(tmpDir, "svc輸出", "Svc端到端", "服務測試範本.odt"));
            Check(body.Contains("12345678") && body.Contains("台中市") && body.Contains("端到端縣政府"),
                "Service 新範本替換值正確", null);
        }

        /// <summary>合成一份含紅/藍彩色 span 的最小 ODT（含相鄰多段紅 span 供合併測試）。</summary>
        static byte[] BuildColoredOdt()
        {
            XNamespace o = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
            XNamespace t = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
            XNamespace s = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
            XNamespace fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

            Func<string, string, XElement> colorStyle = delegate(string name, string hex)
            {
                return new XElement(s + "style",
                    new XAttribute(s + "name", name), new XAttribute(s + "family", "text"),
                    new XElement(s + "text-properties", new XAttribute(fo + "color", hex)));
            };
            Func<string, string, XElement> span = delegate(string style, string text)
            {
                return new XElement(t + "span", new XAttribute(t + "style-name", style), text);
            };

            XDocument doc = new XDocument(new XDeclaration("1.0", "UTF-8", null),
                new XElement(o + "document-content",
                    new XElement(o + "automatic-styles",
                        colorStyle("C_RED", "#FF0000"),
                        colorStyle("C_BLUE", "#0070C0")),
                    new XElement(o + "body",
                        new XElement(o + "text",
                            new XElement(t + "p",
                                span("C_RED", "測試公司股份有限公司"),
                                "（廠商名稱）地址：",
                                span("C_RED", "台北市中正路"), span("C_RED", "1"), span("C_RED", "號"),
                                "，委派至", span("C_BLUE", "某某縣政府"), "辦理。")))));

            byte[] content;
            using (MemoryStream ms = new MemoryStream())
            {
                XmlWriterSettings set = new XmlWriterSettings();
                set.Encoding = new UTF8Encoding(false); set.Indent = false;
                using (XmlWriter xw = XmlWriter.Create(ms, set)) { doc.Save(xw); }
                content = ms.ToArray();
            }
            List<KeyValuePair<string, byte[]>> entries = new List<KeyValuePair<string, byte[]>>();
            entries.Add(new KeyValuePair<string, byte[]>("mimetype",
                Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.text")));
            entries.Add(new KeyValuePair<string, byte[]>("content.xml", content));
            return OdtWriter.Repack(entries);
        }

        // ============================== 工具 ==============================

        static string ReadOdtContentXml(string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            using (StreamReader sr = new StreamReader(zip.GetEntry("content.xml").Open(), Encoding.UTF8))
                return sr.ReadToEnd();
        }

        static string ReadOdtBodyText(string path)
        {
            XDocument xdoc = XDocument.Parse(ReadOdtContentXml(path));
            XNamespace nsOffice = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
            XElement body = xdoc.Descendants(nsOffice + "body").First();
            return body.Value;
        }

        /// <summary>
        /// 手工組一個最小 xlsx（inlineStr、部分列/儲存格故意省略 r 屬性、@SCI@ 佔位符
        /// 會改成無 t 的科學記號數字儲存格）— 同時測到 XlsxReader 的多個相容路徑。
        /// </summary>
        static byte[] BuildXlsx(string[][] tendersRows, string[][] companyRows)
        {
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace nsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace nsRel = "http://schemas.openxmlformats.org/package/2006/relationships";

            Func<string[][], bool, XDocument> buildSheet = delegate(string[][] rows, bool omitR)
            {
                XElement sheetData = new XElement(ns + "sheetData");
                for (int r = 0; r < rows.Length; r++)
                {
                    XElement rowEl = new XElement(ns + "row");
                    if (!omitR || r % 2 == 0) rowEl.SetAttributeValue("r", r + 1);
                    for (int c = 0; c < rows[r].Length; c++)
                    {
                        string val = rows[r][c];
                        if (val == "") continue;   // 空儲存格照真實 Excel 省略（跳號測試）
                        XElement cell = new XElement(ns + "c");
                        if (!omitR || c % 3 != 1)
                            cell.SetAttributeValue("r", ColRef(c) + (r + 1));
                        if (val == "@SCI@")
                        {
                            cell.Add(new XElement(ns + "v", "9.12345678E8"));   // 無 t = 數字
                        }
                        else
                        {
                            cell.SetAttributeValue("t", "inlineStr");
                            cell.Add(new XElement(ns + "is",
                                new XElement(ns + "t",
                                    new XAttribute(XNamespace.Xml + "space", "preserve"), val)));
                        }
                        rowEl.Add(cell);
                    }
                    sheetData.Add(rowEl);
                }
                return new XDocument(new XElement(ns + "worksheet", sheetData));
            };

            XDocument workbook = new XDocument(
                new XElement(ns + "workbook",
                    new XAttribute(XNamespace.Xmlns + "r", nsR),
                    new XElement(ns + "sheets",
                        new XElement(ns + "sheet", new XAttribute("name", "標案清單"),
                            new XAttribute("sheetId", 1), new XAttribute(nsR + "id", "rId1")),
                        new XElement(ns + "sheet", new XAttribute("name", "公司資料"),
                            new XAttribute("sheetId", 2), new XAttribute(nsR + "id", "rId2")))));

            XDocument wbRels = new XDocument(
                new XElement(nsRel + "Relationships",
                    new XElement(nsRel + "Relationship", new XAttribute("Id", "rId1"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                        new XAttribute("Target", "worksheets/sheet1.xml")),
                    new XElement(nsRel + "Relationship", new XAttribute("Id", "rId2"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                        new XAttribute("Target", "worksheets/sheet2.xml"))));

            using (MemoryStream ms = new MemoryStream())
            {
                using (ZipArchive zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
                {
                    WriteXml(zip, "xl/workbook.xml", workbook);
                    WriteXml(zip, "xl/_rels/workbook.xml.rels", wbRels);
                    WriteXml(zip, "xl/worksheets/sheet1.xml", buildSheet(tendersRows, true));
                    WriteXml(zip, "xl/worksheets/sheet2.xml", buildSheet(companyRows, false));
                }
                return ms.ToArray();
            }
        }

        /// <summary>組一個使用 sharedStrings（含傳入的 si）的最小 xlsx，單一儲存格 A1 指向該共享字串。</summary>
        static byte[] BuildXlsxWithSharedString(XElement si)
        {
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace nsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace nsRel = "http://schemas.openxmlformats.org/package/2006/relationships";

            XDocument workbook = new XDocument(
                new XElement(ns + "workbook",
                    new XAttribute(XNamespace.Xmlns + "r", nsR),
                    new XElement(ns + "sheets",
                        new XElement(ns + "sheet", new XAttribute("name", "工作表1"),
                            new XAttribute("sheetId", 1), new XAttribute(nsR + "id", "rId1")))));
            XDocument wbRels = new XDocument(
                new XElement(nsRel + "Relationships",
                    new XElement(nsRel + "Relationship", new XAttribute("Id", "rId1"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                        new XAttribute("Target", "worksheets/sheet1.xml"))));
            XDocument sst = new XDocument(new XElement(ns + "sst", si));
            XDocument sheet = new XDocument(
                new XElement(ns + "worksheet",
                    new XElement(ns + "sheetData",
                        new XElement(ns + "row", new XAttribute("r", 1),
                            new XElement(ns + "c", new XAttribute("r", "A1"), new XAttribute("t", "s"),
                                new XElement(ns + "v", "0"))))));

            using (MemoryStream ms = new MemoryStream())
            {
                using (ZipArchive zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
                {
                    WriteXml(zip, "xl/workbook.xml", workbook);
                    WriteXml(zip, "xl/_rels/workbook.xml.rels", wbRels);
                    WriteXml(zip, "xl/sharedStrings.xml", sst);
                    WriteXml(zip, "xl/worksheets/sheet1.xml", sheet);
                }
                return ms.ToArray();
            }
        }

        static void WriteXml(ZipArchive zip, string name, XDocument doc)
        {
            ZipArchiveEntry entry = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (Stream s = entry.Open())
            using (StreamWriter w = new StreamWriter(s, new UTF8Encoding(false)))
                doc.Save(w, SaveOptions.DisableFormatting);
        }

        static string ColRef(int index)
        {
            string s = "";
            int n = index + 1;
            while (n > 0) { int r = (n - 1) % 26; s = (char)('A' + r) + s; n = (n - 1) / 26; }
            return s;
        }
    }
}
