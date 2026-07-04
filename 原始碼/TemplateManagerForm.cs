using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace TenderDocGen
{
    /// <summary>範本管理：檢視現有範本、新增範本、移除範本（同步更新 Excel）。</summary>
    class TemplateManagerForm : Form
    {
        readonly string _baseDir;
        DataGridView _grid;
        Button _btnAdd, _btnRemove, _btnClose;
        Label _lblStatus;

        /// <summary>是否有變更（新增/移除），供 MainForm 決定是否重整。</summary>
        public bool Changed { get; private set; }

        public TemplateManagerForm(string baseDir)
        {
            _baseDir = baseDir;
            Text = "範本管理";
            Font = new Font("Microsoft JhengHei UI", 12f);
            AutoScaleMode = AutoScaleMode.Font;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(820, 520);
            MinimumSize = new Size(640, 400);

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 1;
            layout.RowCount = 3;
            layout.Padding = new Padding(12);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(layout);

            Label hint = new Label();
            hint.Text = "目前的文件範本。可新增新範本或移除不需要的範本（會同步調整 Excel 欄位）。";
            hint.AutoSize = true; hint.Margin = new Padding(0, 0, 0, 8);
            layout.Controls.Add(hint, 0, 0);

            _grid = new DataGridView();
            _grid.Dock = DockStyle.Fill;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.RowHeadersVisible = false;
            _grid.ReadOnly = true;
            _grid.MultiSelect = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            _grid.BackgroundColor = SystemColors.Window;
            _grid.RowTemplate.MinimumHeight = Font.Height + 14;
            DataGridViewTextBoxColumn cName = new DataGridViewTextBoxColumn();
            cName.HeaderText = "範本名稱"; cName.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; cName.FillWeight = 40;
            DataGridViewTextBoxColumn cTok = new DataGridViewTextBoxColumn();
            cTok.HeaderText = "使用的參數"; cTok.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; cTok.FillWeight = 60;
            cTok.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            _grid.Columns.AddRange(new DataGridViewColumn[] { cName, cTok });
            layout.Controls.Add(_grid, 0, 1);

            FlowLayoutPanel buttons = new FlowLayoutPanel();
            buttons.Dock = DockStyle.Fill; buttons.AutoSize = true;
            _btnAdd = new Button();
            _btnAdd.Text = "新增範本…"; _btnAdd.AutoSize = true; _btnAdd.Padding = new Padding(12, 5, 12, 5);
            _btnAdd.Click += delegate { DoAdd(); };
            _btnRemove = new Button();
            _btnRemove.Text = "移除選取範本"; _btnRemove.AutoSize = true; _btnRemove.Padding = new Padding(12, 5, 12, 5);
            _btnRemove.Margin = new Padding(10, 3, 0, 3);
            _btnRemove.Click += delegate { DoRemove(); };
            _btnClose = new Button();
            _btnClose.Text = "關閉"; _btnClose.AutoSize = true; _btnClose.Padding = new Padding(12, 5, 12, 5);
            _btnClose.Margin = new Padding(10, 3, 0, 3);
            _btnClose.Click += delegate { Close(); };
            _lblStatus = new Label();
            _lblStatus.AutoSize = true; _lblStatus.Margin = new Padding(16, 12, 0, 0);
            buttons.Controls.Add(_btnAdd);
            buttons.Controls.Add(_btnRemove);
            buttons.Controls.Add(_btnClose);
            buttons.Controls.Add(_lblStatus);
            layout.Controls.Add(buttons, 0, 2);

            Shown += delegate { RefreshList(); };
        }

        void RefreshList()
        {
            _grid.Rows.Clear();
            try
            {
                TemplateStore store = TemplateStore.Load(AddTemplateService.TemplateDir(_baseDir));
                foreach (TemplateInfo t in store.Templates)
                {
                    string toks = string.Join("、", t.Tokens.OrderBy(x => x));
                    int idx = _grid.Rows.Add(t.BaseName, toks);
                    _grid.Rows[idx].Tag = t.BaseName;
                }
                _lblStatus.Text = "共 " + store.Templates.Count + " 份範本。";
            }
            catch (Exception ex)
            {
                _lblStatus.Text = "無法讀取範本：" + ex.Message;
            }
        }

        void DoAdd()
        {
            using (AddTemplateForm f = new AddTemplateForm(_baseDir))
            {
                if (f.ShowDialog(this) == DialogResult.OK)
                {
                    Changed = true;
                    RefreshList();
                }
            }
        }

        void DoRemove()
        {
            if (_grid.CurrentRow == null || !(_grid.CurrentRow.Tag is string))
            {
                MessageBox.Show(this, "請先選取要移除的範本。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string baseName = (string)_grid.CurrentRow.Tag;

            RemovalPlan plan = AddTemplateService.PreviewRemoval(baseName, _baseDir);
            if (plan.Error != null)
            {
                MessageBox.Show(this, plan.Error, "無法移除", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            List<string> lines = new List<string> { "即將移除範本「" + baseName + "」，並刪除下列 Excel 欄位：" };
            if (plan.TenderColumns.Count > 0) lines.Add("・標案清單欄：" + string.Join("、", plan.TenderColumns));
            if (plan.CompanyParams.Count > 0) lines.Add("・公司資料參數：" + string.Join("、", plan.CompanyParams));
            lines.Add("");
            lines.Add("這會一併刪除這些欄位／參數已填的資料，且無法復原。確定移除嗎？");
            DialogResult dr = MessageBox.Show(this, string.Join("\n", lines), "確認移除範本",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dr != DialogResult.Yes) return;

            // 需要更新 Excel → 先確保 Excel 關閉
            string xlsx = AddTemplateService.XlsxPath(_baseDir);
            if ((plan.TenderColumns.Count > 0 || plan.CompanyParams.Count > 0) &&
                !ExcelGuard.EnsureClosed(xlsx, this))
                return;

            Cursor = Cursors.WaitCursor;
            _btnRemove.Enabled = _btnAdd.Enabled = false;
            RemovalPlan res;
            try { res = AddTemplateService.RemoveTemplate(baseName, _baseDir); }
            finally { Cursor = Cursors.Default; _btnRemove.Enabled = _btnAdd.Enabled = true; }

            if (!res.Ok)
            {
                MessageBox.Show(this, "移除失敗：\n" + res.Error, "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Changed = true;
            RefreshList();
            MessageBox.Show(this, "已移除範本「" + baseName + "」。", "完成",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
