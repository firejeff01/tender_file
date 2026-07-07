using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace TenderDocGen
{
    /// <summary>「新增範本」對話框：選彩色 ODT → 對應參數 → 產生範本並更新 Excel。</summary>
    class AddTemplateForm : Form
    {
        const string Pick = "（請選擇）";
        const string Skip = "（略過此段）";
        const string AddNew = "＋ 新增參數…";

        readonly string _baseDir;
        string _odtPath;
        string _tempOdt;               // Word 轉入時的暫存 ODT（關閉/重選時清除）
        TemplateBuilder _builder;
        List<ColoredRun> _runs;
        readonly Dictionary<string, bool> _newTokenIsCompany = new Dictionary<string, bool>();
        Dictionary<string, string> _valueToToken = new Dictionary<string, string>();  // 現有資料值→token（預填）

        Button _btnPick, _btnOk, _btnCancel;
        Label _lblFile;
        TextBox _txtName;
        DataGridView _grid;
        DataGridViewComboBoxColumn _colToken;

        const int CText = 0, CColor = 1, CToken = 2;

        public AddTemplateForm(string baseDir)
        {
            _baseDir = baseDir;
            LoadPrefillValues();

            Text = "新增範本";
            UiTheme.StyleForm(this);
            AutoScaleMode = AutoScaleMode.None;
            StartPosition = FormStartPosition.CenterParent;
            UiTheme.SizeWindow(this, 1060, 740, 820, 560);

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.BackColor = UiTheme.BgMain;
            layout.ColumnCount = 1;
            layout.RowCount = 4;
            layout.Padding = new Padding(18, 12, 18, 12);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(layout);

            // 第 0 列：選檔
            FlowLayoutPanel pick = new FlowLayoutPanel();
            pick.Dock = DockStyle.Fill; pick.AutoSize = true; pick.WrapContents = true;
            pick.Margin = new Padding(0, 0, 0, 6);
            _btnPick = new Button();
            _btnPick.Text = "📄 選擇文件…"; _btnPick.AutoSize = true; _btnPick.Padding = new Padding(14, 7, 14, 7);
            UiTheme.StylePrimary(_btnPick);
            _btnPick.Click += delegate { PickFile(); };
            _lblFile = new Label();
            _lblFile.Text = "請選擇一份「已把要替換的文字標成顏色」的 ODT 或 Word 文件。";
            _lblFile.AutoSize = true; _lblFile.Margin = new Padding(12, 12, 0, 0);
            UiTheme.StyleHint(_lblFile);
            pick.Controls.Add(_btnPick); pick.Controls.Add(_lblFile);
            layout.Controls.Add(pick, 0, 0);

            // 第 1 列：範本名稱
            FlowLayoutPanel nameRow = new FlowLayoutPanel();
            nameRow.Dock = DockStyle.Fill; nameRow.AutoSize = true; nameRow.Margin = new Padding(0, 0, 0, 6);
            Label lblN = new Label();
            lblN.Text = "範本名稱"; lblN.AutoSize = true; lblN.Margin = new Padding(0, 9, 10, 0);
            lblN.ForeColor = UiTheme.TextSecondary;
            _txtName = new TextBox();
            _txtName.Width = 460; _txtName.Margin = new Padding(0, 4, 0, 0);
            UiTheme.StyleTextBox(_txtName);
            nameRow.Controls.Add(lblN); nameRow.Controls.Add(_txtName);
            layout.Controls.Add(nameRow, 0, 1);

            // 第 2 列：對應表
            _grid = new DataGridView();
            _grid.Dock = DockStyle.Fill;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.RowHeadersVisible = false;
            _grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            _grid.EditMode = DataGridViewEditMode.EditOnEnter;
            UiTheme.StyleGrid(_grid);

            DataGridViewTextBoxColumn cText = new DataGridViewTextBoxColumn();
            cText.HeaderText = "文字片段"; cText.ReadOnly = true;
            cText.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; cText.FillWeight = 55;
            cText.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            DataGridViewTextBoxColumn cColor = new DataGridViewTextBoxColumn();
            cColor.HeaderText = "顏色"; cColor.ReadOnly = true;
            cColor.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            _colToken = new DataGridViewComboBoxColumn();
            _colToken.HeaderText = "對應參數";
            _colToken.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; _colToken.FillWeight = 45;
            _colToken.FlatStyle = FlatStyle.Flat;
            foreach (string it in TokenChoices()) _colToken.Items.Add(it);
            _grid.Columns.AddRange(new DataGridViewColumn[] { cText, cColor, _colToken });
            _grid.CurrentCellDirtyStateChanged += delegate
            {
                if (_grid.IsCurrentCellDirty) _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            _grid.CellValueChanged += Grid_CellValueChanged;
            _grid.DataError += delegate { };   // 吞掉 combо 短暫不一致的 DataError
            layout.Controls.Add(_grid, 0, 2);

            // 第 3 列：確定/取消
            FlowLayoutPanel buttons = new FlowLayoutPanel();
            buttons.Dock = DockStyle.Fill; buttons.FlowDirection = FlowDirection.RightToLeft;
            buttons.AutoSize = true; buttons.Margin = new Padding(0, 8, 0, 0);
            _btnOk = new Button();
            _btnOk.Text = "✔ 建立範本"; _btnOk.AutoSize = true; _btnOk.Padding = new Padding(18, 8, 18, 8);
            _btnOk.Enabled = false;
            UiTheme.StyleSuccess(_btnOk);
            _btnOk.Click += delegate { DoCommit(); };
            _btnCancel = new Button();
            _btnCancel.Text = "取消"; _btnCancel.AutoSize = true; _btnCancel.Padding = new Padding(16, 8, 16, 8);
            _btnCancel.Margin = new Padding(0, 0, 12, 0);
            UiTheme.StyleButton(_btnCancel);
            _btnCancel.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
            buttons.Controls.Add(_btnOk); buttons.Controls.Add(_btnCancel);
            layout.Controls.Add(buttons, 0, 3);

            // 頁首（Top，最後加入）
            Controls.Add(UiTheme.Header("新增範本", null));
        }

        List<string> TokenChoices()
        {
            List<string> list = new List<string> { Pick, Skip };
            list.AddRange(AddTemplateService.ExistingTokens);
            foreach (string t in _newTokenIsCompany.Keys) if (!list.Contains(t)) list.Add(t);
            list.Add(AddNew);
            return list;
        }

        void LoadPrefillValues()
        {
            // 用目前「公司資料」的值做預填：文件裡若出現公司名/電話等，直接對上參數
            try
            {
                string xlsx = AddTemplateService.XlsxPath(_baseDir);
                if (!File.Exists(xlsx)) return;
                Dictionary<string, List<XlsxRow>> book = XlsxReader.Load(xlsx);
                foreach (KeyValuePair<string, List<XlsxRow>> kv in book)
                {
                    if (Util.Nfkc(kv.Key) != "公司資料") continue;
                    foreach (XlsxRow r in kv.Value)
                    {
                        string key = Util.Nfkc(r.Cell(0));
                        string val = (r.Cell(1) ?? "").Trim();
                        if (val != "" && AddTemplateService.CompanyTokens.Contains(key))
                            _valueToToken[val] = key;
                    }
                }
            }
            catch { }
            // 加上原始開發範本的示範值（沿用原本規則），能對上就零操作
            AddSample("哈瑪星科技股份有限公司", "廠商名稱");
            AddSample("哈瑪星科技有限公司", "廠商名稱");
            AddSample("陳文德", "負責人姓名");
            AddSample("(07)536-4800", "聯絡電話");
            AddSample("花蓮縣政府", "機關名稱");
        }

        void AddSample(string value, string token)
        {
            if (!_valueToToken.ContainsKey(value)) _valueToToken[value] = token;
        }

        void PickFile()
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "文件範本 (*.odt;*.docx)|*.odt;*.docx|ODF 文字文件 (*.odt)|*.odt|" +
                             "Word 文件 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                dlg.Title = "選擇要作為範本的文件（ODT 或 Word）";
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                LoadSource(dlg.FileName);
            }
        }

        // 選檔入口：Word 先轉成暫存 ODT 再走既有流程；ODT 直接載入。
        void LoadSource(string path)
        {
            _btnOk.Enabled = false;
            _odtPath = null;
            CleanupTemp();
            // 重選檔案時先清空舊狀態，避免載入失敗後仍顯示上一份檔案的對應表（誤導）
            _grid.Rows.Clear();
            _txtName.Text = "";
            _lblFile.Text = "請選擇一份「已把要替換的文字標成顏色」的 ODT 或 Word 文件。";

            string odtPath = path;
            if (string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
            {
                byte[] odtBytes;
                List<string> warnings;
                try
                {
                    odtBytes = DocxToOdt.Convert(path, out warnings);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "無法把這份 Word 轉為範本：\n\n" + ex.Message, "無法轉換 Word",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                try
                {
                    _tempOdt = Path.Combine(Path.GetTempPath(), "tdg_" + Guid.NewGuid().ToString("N") + ".odt");
                    File.WriteAllBytes(_tempOdt, odtBytes);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "無法建立暫存檔：\n" + ex.Message, "錯誤",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                odtPath = _tempOdt;
                if (warnings != null && warnings.Count > 0)
                    MessageBox.Show(this, "已從 Word 轉入，但有以下提醒：\n・" + string.Join("\n・", warnings),
                        "轉換提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            LoadOdt(odtPath, Path.GetFileName(path), Path.GetFileNameWithoutExtension(path));
        }

        void LoadOdt(string odtPath, string displayFile, string displayName)
        {
            try
            {
                _builder = TemplateBuilder.Load(odtPath);
                _runs = _builder.ExtractRuns();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "無法讀取這份文件：\n" + ex.Message, "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (_runs.Count == 0)
            {
                MessageBox.Show(this,
                    "這份文件裡偵測不到彩色文字。\n請把要替換的欄位用「字型色彩（文字顏色）」標成顏色"
                    + "（例如公司資料用紅色、每案資料用藍色）再試一次。\n"
                    + "註：螢光筆／醒目提示、或套用「字元樣式」帶出的顏色不會被偵測到，請直接改文字顏色。",
                    "沒有可參數化的內容", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            _odtPath = odtPath;
            _lblFile.Text = displayFile + "（偵測到 " + _runs.Count + " 段彩色文字）";
            _txtName.Text = displayName;

            _grid.Rows.Clear();
            foreach (ColoredRun run in _runs)
            {
                int idx = _grid.Rows.Add(run.Text, ColorLabel(run), SuggestToken(run));
                DataGridViewRow row = _grid.Rows[idx];
                row.Tag = run;
                Color swatch = HexToColor(run.ColorHex);
                row.Cells[CColor].Style.BackColor = swatch;
                row.Cells[CColor].Style.ForeColor = Brightness(swatch) < 130 ? Color.White : Color.Black;
            }
            _btnOk.Enabled = true;
        }

        string SuggestToken(ColoredRun run)
        {
            string t = run.Text.Trim();
            string token;
            if (_valueToToken.TryGetValue(t, out token)) return token;
            return Pick;
        }

        void CleanupTemp()
        {
            if (_tempOdt == null) return;
            try { if (File.Exists(_tempOdt)) File.Delete(_tempOdt); }
            catch { /* 暫存檔清不掉不影響結果，忽略 */ }
            _tempOdt = null;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            CleanupTemp();
            base.OnFormClosed(e);
        }

        void Grid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != CToken || e.RowIndex < 0) return;
            DataGridViewCell cell = _grid.Rows[e.RowIndex].Cells[CToken];
            if ((string)cell.Value != AddNew) return;

            bool isCompany;
            string name = PromptNewToken(this, out isCompany);
            if (name == null)
            {
                cell.Value = Pick;   // 取消
                return;
            }
            if (!_newTokenIsCompany.ContainsKey(name) && !AddTemplateService.IsExisting(name))
            {
                _newTokenIsCompany[name] = isCompany;
                // 補進所有下拉選項（插在「新增參數」前）
                if (!_colToken.Items.Contains(name))
                    _colToken.Items.Insert(_colToken.Items.Count - 1, name);
            }
            cell.Value = name;
        }

        void DoCommit()
        {
            if (_odtPath == null) return;

            // 收集對應
            Dictionary<string, string> map = new Dictionary<string, string>();
            int unresolved = 0;
            foreach (DataGridViewRow row in _grid.Rows)
            {
                ColoredRun run = row.Tag as ColoredRun;
                if (run == null) continue;
                string v = (string)row.Cells[CToken].Value;
                if (v == Pick) { unresolved++; continue; }
                if (v == Skip || v == AddNew || string.IsNullOrEmpty(v)) continue;
                map[run.Id] = v;
            }

            if (map.Count == 0)
            {
                MessageBox.Show(this, "尚未指定任何參數對應。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (unresolved > 0)
            {
                DialogResult dr = MessageBox.Show(this,
                    "還有 " + unresolved + " 段標成「（請選擇）」尚未指定，這些會被略過（保留原樣）。\n要繼續嗎？",
                    "尚有未指定", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr != DialogResult.Yes) return;
            }

            bool overwrite = false;
            string name = _txtName.Text.Trim();
            string target = Path.Combine(AddTemplateService.TemplateDir(_baseDir), name + ".odt");
            if (File.Exists(target))
            {
                DialogResult dr = MessageBox.Show(this,
                    "已存在同名範本「" + name + "」，要覆蓋嗎？", "確認覆蓋",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dr != DialogResult.Yes) return;
                overwrite = true;
            }

            // 需要更新 Excel → 先確保 Excel 已關閉
            if (!ExcelGuard.EnsureClosed(AddTemplateService.XlsxPath(_baseDir), this)) return;

            Cursor = Cursors.WaitCursor;
            _btnOk.Enabled = false;
            AddTemplateResult res;
            try
            {
                res = AddTemplateService.Commit(_odtPath, name, map, _newTokenIsCompany, _baseDir, overwrite);
            }
            finally { Cursor = Cursors.Default; _btnOk.Enabled = true; }

            if (!res.Ok)
            {
                MessageBox.Show(this, "建立失敗：\n" + res.Error, "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            List<string> lines = new List<string> { "已建立範本：" + res.TemplateFile };
            List<string> newCols = res.AddedTenderColumns.Where(c => !c.StartsWith("產生:")).ToList();
            if (newCols.Count > 0) lines.Add("標案清單新增欄位：" + string.Join("、", newCols));
            if (res.AddedCompanyParams.Count > 0)
                lines.Add("公司資料新增參數：" + string.Join("、", res.AddedCompanyParams) + "（記得填入固定值）");
            lines.Add("已在標案清單加入「產生:" + name + "」欄。");
            MessageBox.Show(this, string.Join("\n", lines), "完成",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            DialogResult = DialogResult.OK;
            Close();
        }

        // ---------- 小工具 ----------

        static string ColorLabel(ColoredRun run)
        {
            if (run.Category == "red") return "紅";
            if (run.Category == "blue") return "藍";
            return run.ColorHex;
        }

        static Color HexToColor(string hex)
        {
            try { return ColorTranslator.FromHtml(hex); }
            catch { return Color.White; }
        }

        static int Brightness(Color c) { return (c.R * 299 + c.G * 587 + c.B * 114) / 1000; }

        /// <summary>小型輸入框：新參數名稱 + 類型（每案／公司固定）。取消回傳 null。</summary>
        static string PromptNewToken(IWin32Window owner, out bool isCompany)
        {
            isCompany = false;
            using (Form f = new Form())
            {
                f.Text = "新增參數";
                UiTheme.StyleForm(f);
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.StartPosition = FormStartPosition.CenterParent;
                f.MinimizeBox = false; f.MaximizeBox = false; f.ShowInTaskbar = false;
                // 版面 AutoSize（隨字型/DPI 自適應，不用固定像素座標）
                f.AutoSize = true;
                f.AutoSizeMode = AutoSizeMode.GrowAndShrink;

                TableLayoutPanel tp = new TableLayoutPanel();
                tp.Dock = DockStyle.Fill; tp.AutoSize = true; tp.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                tp.ColumnCount = 1; tp.RowCount = 4;
                tp.Padding = new Padding(22, 18, 22, 16);
                tp.BackColor = UiTheme.BgMain;
                tp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
                for (int i = 0; i < 4; i++) tp.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                Label lbl = new Label();
                lbl.Text = "參數名稱（會成為 Excel 欄位名）"; lbl.AutoSize = true;
                lbl.ForeColor = UiTheme.TextSecondary; lbl.Margin = new Padding(0, 0, 0, 6);

                TextBox txt = new TextBox();
                txt.Dock = DockStyle.Fill; txt.MinimumSize = new Size(360, 0);
                UiTheme.StyleTextBox(txt); txt.Margin = new Padding(0, 0, 0, 16);

                GroupBox gb = new GroupBox();
                gb.Text = "這個參數的類型"; gb.AutoSize = true; gb.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                gb.ForeColor = UiTheme.TextPrimary; gb.Dock = DockStyle.Fill; gb.Margin = new Padding(0, 0, 0, 16);
                FlowLayoutPanel radios = new FlowLayoutPanel();
                radios.Dock = DockStyle.Fill; radios.FlowDirection = FlowDirection.TopDown;
                radios.AutoSize = true; radios.WrapContents = false; radios.Padding = new Padding(12, 8, 12, 10);
                RadioButton rbCase = new RadioButton();
                rbCase.Text = "每案不同（填在「標案清單」每一列）"; rbCase.AutoSize = true; rbCase.Checked = true;
                rbCase.Margin = new Padding(0, 4, 0, 4); rbCase.ForeColor = UiTheme.TextPrimary;
                RadioButton rbCompany = new RadioButton();
                rbCompany.Text = "公司固定（填一次在「公司資料」）"; rbCompany.AutoSize = true;
                rbCompany.Margin = new Padding(0, 4, 0, 4); rbCompany.ForeColor = UiTheme.TextPrimary;
                radios.Controls.Add(rbCase); radios.Controls.Add(rbCompany);
                gb.Controls.Add(radios);

                FlowLayoutPanel btns = new FlowLayoutPanel();
                btns.Dock = DockStyle.Fill; btns.FlowDirection = FlowDirection.RightToLeft; btns.AutoSize = true;
                Button ok = new Button();
                ok.Text = "確定"; ok.DialogResult = DialogResult.OK; ok.AutoSize = true;
                ok.Padding = new Padding(22, 8, 22, 8); UiTheme.StylePrimary(ok);
                Button cancel = new Button();
                cancel.Text = "取消"; cancel.DialogResult = DialogResult.Cancel; cancel.AutoSize = true;
                cancel.Padding = new Padding(18, 8, 18, 8); cancel.Margin = new Padding(0, 0, 12, 0);
                UiTheme.StyleButton(cancel);
                btns.Controls.Add(ok); btns.Controls.Add(cancel);

                tp.Controls.Add(lbl, 0, 0);
                tp.Controls.Add(txt, 0, 1);
                tp.Controls.Add(gb, 0, 2);
                tp.Controls.Add(btns, 0, 3);
                f.Controls.Add(tp);
                f.AcceptButton = ok; f.CancelButton = cancel;

                if (f.ShowDialog(owner) != DialogResult.OK) return null;
                string name = txt.Text.Trim();
                if (name == "") return null;
                if (!AddTemplateService.IsValidTokenName(name))
                {
                    MessageBox.Show(owner, "參數名稱不可包含 { } $ 或控制字元，請改用其他名稱。",
                        "名稱不合法", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }
                isCompany = rbCompany.Checked;
                return name;
            }
        }
    }
}
