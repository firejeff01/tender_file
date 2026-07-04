using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TenderDocGen
{
    /// <summary>主視窗：列出 Excel 中的標案、勾選、產生文件、回報結果。</summary>
    class MainForm : Form
    {
        readonly string _baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string XlsxPath { get { return Path.Combine(_baseDir, "標案資料.xlsx"); } }
        string TemplateDir { get { return Path.Combine(_baseDir, "範本"); } }
        string OutputRoot { get { return _outputRoot; } }

        Settings _settings;
        string _outputRoot;

        DataGridView _grid;
        Button _btnReload, _btnGenerate, _btnOpenOutput, _btnChangeOutput, _btnAddTemplate;
        CheckBox _chkOverwrite;
        Label _lblStatus;
        TextBox _txtOutput;

        TemplateStore _store;
        GenerationPlan _plan;

        const int ColCheck = 0, ColRowNo = 1, ColName = 2, ColDocs = 3, ColState = 4, ColMessage = 5;

        public MainForm()
        {
            _settings = new Settings(_baseDir);
            InitOutputRoot();

            Text = "標案文件產生器";
            // 加大字級提升可讀性；AutoScaleMode=Font 會讓整個介面等比放大
            Font = new Font("Microsoft JhengHei UI", 12f);
            AutoScaleMode = AutoScaleMode.Font;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(1180, 720);
            MinimumSize = new Size(900, 560);

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 1;
            layout.RowCount = 4;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // 說明＋重新讀取
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // 輸出資料夾設定
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f)); // 清單
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // 操作列
            layout.Padding = new Padding(12);
            Controls.Add(layout);

            // --- 第 0 列：說明與重新讀取 ---
            FlowLayoutPanel top = new FlowLayoutPanel();
            top.Dock = DockStyle.Fill;
            top.AutoSize = true;
            top.WrapContents = true;
            top.Margin = new Padding(0, 0, 0, 6);
            _btnReload = new Button();
            _btnReload.Text = "重新讀取 Excel";
            _btnReload.AutoSize = true;
            _btnReload.Padding = new Padding(8, 4, 8, 4);
            _btnReload.Click += delegate { LoadPlan(); };
            _btnAddTemplate = new Button();
            _btnAddTemplate.Text = "範本管理…";
            _btnAddTemplate.AutoSize = true;
            _btnAddTemplate.Padding = new Padding(8, 4, 8, 4);
            _btnAddTemplate.Margin = new Padding(10, 3, 0, 3);
            _btnAddTemplate.Click += delegate { ManageTemplates(); };
            Label hint = new Label();
            hint.Text = "在「標案資料.xlsx」填寫標案並存檔後，按此重新讀取。";
            hint.AutoSize = true;
            hint.Margin = new Padding(10, 10, 0, 0);
            top.Controls.Add(_btnReload);
            top.Controls.Add(_btnAddTemplate);
            top.Controls.Add(hint);
            layout.Controls.Add(top, 0, 0);

            // --- 第 1 列：輸出資料夾設定 ---
            TableLayoutPanel outRow = new TableLayoutPanel();
            outRow.Dock = DockStyle.Fill;
            outRow.AutoSize = true;
            outRow.ColumnCount = 3;
            outRow.RowCount = 1;
            outRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            outRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            outRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            outRow.Margin = new Padding(0, 0, 0, 8);
            Label lblOut = new Label();
            lblOut.Text = "輸出資料夾：";
            lblOut.AutoSize = true;
            lblOut.Anchor = AnchorStyles.Left;
            lblOut.Margin = new Padding(0, 8, 4, 0);
            _txtOutput = new TextBox();
            _txtOutput.ReadOnly = true;
            _txtOutput.Dock = DockStyle.Fill;
            _txtOutput.Text = _outputRoot;
            _txtOutput.Margin = new Padding(0, 4, 6, 0);
            _btnChangeOutput = new Button();
            _btnChangeOutput.Text = "變更…";
            _btnChangeOutput.AutoSize = true;
            _btnChangeOutput.Padding = new Padding(8, 2, 8, 2);
            _btnChangeOutput.Click += delegate { ChangeOutputFolder(); };
            outRow.Controls.Add(lblOut, 0, 0);
            outRow.Controls.Add(_txtOutput, 1, 0);
            outRow.Controls.Add(_btnChangeOutput, 2, 0);
            layout.Controls.Add(outRow, 0, 1);

            // --- 第 2 列：清單 ---
            _grid = new DataGridView();
            _grid.Dock = DockStyle.Fill;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.AllowUserToResizeRows = false;
            _grid.RowHeadersVisible = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect = false;
            _grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            _grid.ColumnHeadersDefaultCellStyle.Padding = new Padding(4, 8, 4, 8);
            _grid.EnableHeadersVisualStyles = true;
            _grid.BackgroundColor = SystemColors.Window;
            _grid.BorderStyle = BorderStyle.Fixed3D;
            _grid.RowTemplate.MinimumHeight = Font.Height + 16;   // 加高行高，閱讀更輕鬆
            _grid.DefaultCellStyle.Padding = new Padding(3, 4, 3, 4);

            DataGridViewCheckBoxColumn colCheck = new DataGridViewCheckBoxColumn();
            colCheck.HeaderText = "勾選";
            colCheck.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            colCheck.Width = (int)(Font.Height * 2.4);
            DataGridViewTextBoxColumn colRowNo = MakeTextCol("列", DataGridViewAutoSizeColumnMode.AllCells);
            DataGridViewTextBoxColumn colName = MakeTextCol("標案名稱", DataGridViewAutoSizeColumnMode.Fill);
            colName.FillWeight = 45;
            DataGridViewTextBoxColumn colDocs = MakeTextCol("文件數", DataGridViewAutoSizeColumnMode.AllCells);
            DataGridViewTextBoxColumn colState = MakeTextCol("狀態", DataGridViewAutoSizeColumnMode.AllCells);
            DataGridViewTextBoxColumn colMsg = MakeTextCol("訊息", DataGridViewAutoSizeColumnMode.Fill);
            colMsg.FillWeight = 55;
            colMsg.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            _grid.Columns.AddRange(new DataGridViewColumn[] {
                colCheck, colRowNo, colName, colDocs, colState, colMsg });
            layout.Controls.Add(_grid, 0, 2);

            // --- 第 3 列：操作列 ---
            FlowLayoutPanel bottom = new FlowLayoutPanel();
            bottom.Dock = DockStyle.Fill;
            bottom.AutoSize = true;
            bottom.WrapContents = true;
            bottom.Margin = new Padding(0, 6, 0, 0);
            _btnGenerate = new Button();
            _btnGenerate.Text = "產生文件";
            _btnGenerate.AutoSize = true;
            _btnGenerate.Padding = new Padding(16, 5, 16, 5);
            _btnGenerate.Font = new Font(Font, FontStyle.Bold);
            _btnGenerate.Click += delegate { GenerateChecked(); };
            _chkOverwrite = new CheckBox();
            _chkOverwrite.Text = "重新產生（覆寫已存在的檔案）";
            _chkOverwrite.AutoSize = true;
            _chkOverwrite.Margin = new Padding(14, 10, 0, 0);
            _btnOpenOutput = new Button();
            _btnOpenOutput.Text = "開啟輸出資料夾";
            _btnOpenOutput.AutoSize = true;
            _btnOpenOutput.Padding = new Padding(8, 4, 8, 4);
            _btnOpenOutput.Margin = new Padding(14, 3, 0, 0);
            _btnOpenOutput.Click += delegate { OpenOutputFolder(); };
            _lblStatus = new Label();
            _lblStatus.AutoSize = true;
            _lblStatus.Margin = new Padding(18, 12, 0, 0);
            bottom.Controls.Add(_btnGenerate);
            bottom.Controls.Add(_chkOverwrite);
            bottom.Controls.Add(_btnOpenOutput);
            bottom.Controls.Add(_lblStatus);
            layout.Controls.Add(bottom, 0, 3);

            Shown += delegate { LoadPlan(); };
        }

        void InitOutputRoot()
        {
            string configured = _settings.Get(Settings.KeyOutputDir, "");
            if (configured == "") _outputRoot = Path.Combine(_baseDir, "輸出");
            else if (Path.IsPathRooted(configured)) _outputRoot = configured;
            else _outputRoot = Path.Combine(_baseDir, configured);
        }

        DataGridViewTextBoxColumn MakeTextCol(string header, DataGridViewAutoSizeColumnMode sizeMode)
        {
            DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
            col.HeaderText = header;
            col.AutoSizeMode = sizeMode;
            col.ReadOnly = true;
            col.SortMode = DataGridViewColumnSortMode.NotSortable;
            return col;
        }

        // ==================== 範本管理 ====================

        void ManageTemplates()
        {
            using (TemplateManagerForm f = new TemplateManagerForm(_baseDir))
            {
                f.ShowDialog(this);
                if (f.Changed) LoadPlan();   // 範本有增減 → 重新掃描範本＋Excel
            }
        }

        // ==================== 輸出資料夾 ====================

        void ChangeOutputFolder()
        {
            using (FolderBrowserDialog dlg = new FolderBrowserDialog())
            {
                dlg.Description = "選擇文件要輸出到哪個資料夾";
                dlg.ShowNewFolderButton = true;
                if (Directory.Exists(_outputRoot)) dlg.SelectedPath = _outputRoot;
                else if (Directory.Exists(_baseDir)) dlg.SelectedPath = _baseDir;
                if (dlg.ShowDialog(this) == DialogResult.OK && dlg.SelectedPath != "")
                {
                    _outputRoot = dlg.SelectedPath;
                    _settings.Set(Settings.KeyOutputDir, _outputRoot);
                    _txtOutput.Text = _outputRoot;
                    LoadPlan();   // 重新整理「已產生 / 尚未產生」狀態（依輸出路徑而定）
                }
            }
        }

        // ==================== 載入 ====================

        void LoadPlan()
        {
            _grid.Rows.Clear();
            _lblStatus.Text = "";
            _plan = null;

            try { _store = TemplateStore.Load(TemplateDir); }
            catch (Exception ex)
            {
                ShowFatal("讀取範本失敗：\n" + ex.Message);
                return;
            }

            try { _plan = Planner.BuildPlan(XlsxPath, _store); }
            catch (Exception ex)
            {
                ShowFatal("讀取 Excel 失敗：\n" + ex.Message);
                return;
            }

            if (_plan.GlobalErrors.Count > 0)
            {
                MessageBox.Show(this, string.Join("\n", _plan.GlobalErrors), "標案資料.xlsx 有問題",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            foreach (RowPlan rp in _plan.Rows)
            {
                string state = RowState(rp);
                bool check = rp.Errors.Count == 0 && rp.Docs.Count > 0 && state != "已產生";
                List<string> msgs = new List<string>();
                msgs.AddRange(rp.Errors.Select(e => "✘ " + e));
                msgs.AddRange(rp.Warnings.Select(w => "⚠ " + w));

                int idx = _grid.Rows.Add(check, rp.RowNumber, rp.TenderName,
                    rp.Docs.Count, state, string.Join("\n", msgs));
                DataGridViewRow row = _grid.Rows[idx];
                row.Tag = rp;
                if (rp.Errors.Count > 0)
                {
                    row.DefaultCellStyle.ForeColor = Color.Firebrick;
                    row.Cells[ColCheck].ReadOnly = true;
                }
            }
            UpdateStatus(string.Format("已讀取 {0} 個標案。", _plan.Rows.Count));
        }

        string RowState(RowPlan rp)
        {
            if (rp.Errors.Count > 0) return "資料有誤";
            if (rp.Docs.Count == 0) return "未選文件";
            if (rp.FolderName == null) return "資料有誤";
            string folder = Path.Combine(OutputRoot, rp.FolderName);
            if (!Directory.Exists(folder)) return "尚未產生";
            int existing = rp.Docs.Count(t => File.Exists(Path.Combine(folder, t.FileName)));
            if (existing == rp.Docs.Count) return "已產生";
            if (existing == 0) return "尚未產生";
            return "部分產生";
        }

        // ==================== 產生 ====================

        void GenerateChecked()
        {
            if (_plan == null) { LoadPlan(); if (_plan == null) return; }

            List<DataGridViewRow> targets = new List<DataGridViewRow>();
            foreach (DataGridViewRow row in _grid.Rows)
                if (Convert.ToBoolean(row.Cells[ColCheck].Value) && row.Tag is RowPlan)
                    targets.Add(row);

            if (targets.Count == 0)
            {
                MessageBox.Show(this, "請先勾選要產生的標案。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool overwrite = _chkOverwrite.Checked;
            _btnGenerate.Enabled = _btnReload.Enabled = _btnChangeOutput.Enabled = false;
            Cursor = Cursors.WaitCursor;
            int okRows = 0, errRows = 0, totalFiles = 0, skippedFiles = 0;
            try
            {
                foreach (DataGridViewRow row in targets)
                {
                    RowPlan rp = (RowPlan)row.Tag;
                    RowResult r = Generator.GenerateRow(rp, OutputRoot, overwrite, _plan.KeepRed);
                    totalFiles += r.Generated.Count;
                    skippedFiles += r.Skipped.Count;

                    List<string> msgs = new List<string>();
                    msgs.AddRange(r.Errors.Select(e => "✘ " + e));
                    msgs.AddRange(rp.Warnings.Select(w => "⚠ " + w));
                    if (r.Generated.Count > 0) msgs.Add("✔ 已產生：" + string.Join("、", r.Generated));
                    if (r.Skipped.Count > 0) msgs.Add("─ 已存在略過：" + string.Join("、", r.Skipped));

                    if (r.Errors.Count > 0)
                    {
                        errRows++;
                        row.Cells[ColState].Value = "有錯誤";
                        row.DefaultCellStyle.ForeColor = Color.Firebrick;
                    }
                    else
                    {
                        okRows++;
                        row.Cells[ColState].Value = RowState(rp);
                        row.Cells[ColCheck].Value = false;
                        row.DefaultCellStyle.ForeColor = Color.ForestGreen;
                    }
                    row.Cells[ColMessage].Value = string.Join("\n", msgs);
                }
            }
            finally
            {
                Cursor = Cursors.Default;
                _btnGenerate.Enabled = _btnReload.Enabled = _btnChangeOutput.Enabled = true;
            }

            UpdateStatus(string.Format("完成：{0} 個標案成功、{1} 個有錯誤；產生 {2} 個檔案、略過 {3} 個。",
                okRows, errRows, totalFiles, skippedFiles));
            if (errRows == 0 && totalFiles > 0)
            {
                DialogResult dr = MessageBox.Show(this,
                    "已產生 " + totalFiles + " 個文件。\n要開啟輸出資料夾嗎？", "完成",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (dr == DialogResult.Yes) OpenOutputFolder();
            }
        }

        void OpenOutputFolder()
        {
            try
            {
                Directory.CreateDirectory(OutputRoot);
                Process.Start("explorer.exe", "\"" + OutputRoot + "\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "無法開啟輸出資料夾：" + ex.Message, "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void UpdateStatus(string text) { _lblStatus.Text = text; }

        void ShowFatal(string message)
        {
            MessageBox.Show(this, message + "\n\n請確認本程式與「標案資料.xlsx」「範本」資料夾放在同一個資料夾內。",
                "無法載入", MessageBoxButtons.OK, MessageBoxIcon.Error);
            UpdateStatus("載入失敗，請修正後按「重新讀取 Excel」。");
        }
    }
}
