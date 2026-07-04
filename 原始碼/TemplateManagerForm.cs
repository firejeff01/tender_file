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
            UiTheme.StyleForm(this);
            AutoScaleMode = AutoScaleMode.Font;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(1000, 660);
            MinimumSize = new Size(760, 480);

            // ===== 清單（Fill）=====
            _grid = new DataGridView();
            _grid.Dock = DockStyle.Fill;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.RowHeadersVisible = false;
            _grid.ReadOnly = true;
            _grid.MultiSelect = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            UiTheme.StyleGrid(_grid);
            DataGridViewTextBoxColumn cName = new DataGridViewTextBoxColumn();
            cName.HeaderText = "範本名稱"; cName.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; cName.FillWeight = 40;
            DataGridViewTextBoxColumn cTok = new DataGridViewTextBoxColumn();
            cTok.HeaderText = "使用的參數"; cTok.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; cTok.FillWeight = 60;
            cTok.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            _grid.Columns.AddRange(new DataGridViewColumn[] { cName, cTok });
            Panel gridWrap = new Panel();
            gridWrap.Dock = DockStyle.Fill; gridWrap.BackColor = UiTheme.BgMain;
            gridWrap.Padding = new Padding(18, 10, 18, 8);
            gridWrap.Controls.Add(_grid);
            Controls.Add(gridWrap);

            // ===== 操作列（Bottom）=====
            Panel footer = new Panel();
            footer.Dock = DockStyle.Bottom; footer.AutoSize = true;
            footer.AutoSizeMode = AutoSizeMode.GrowAndShrink; footer.BackColor = UiTheme.BgCard;
            FlowLayoutPanel buttons = new FlowLayoutPanel();
            buttons.Dock = DockStyle.Top; buttons.AutoSize = true; buttons.WrapContents = false;
            buttons.Padding = new Padding(18, 12, 18, 12);
            _btnAdd = new Button();
            _btnAdd.Text = "＋ 新增範本…"; _btnAdd.AutoSize = true; _btnAdd.Padding = new Padding(16, 8, 16, 8);
            UiTheme.StylePrimary(_btnAdd);
            _btnAdd.Click += delegate { DoAdd(); };
            _btnRemove = new Button();
            _btnRemove.Text = "🗑 移除選取範本"; _btnRemove.AutoSize = true; _btnRemove.Padding = new Padding(14, 8, 14, 8);
            _btnRemove.Anchor = AnchorStyles.Left; _btnRemove.Margin = new Padding(12, 0, 0, 0);
            UiTheme.StyleDanger(_btnRemove);
            _btnRemove.Click += delegate { DoRemove(); };
            _btnClose = new Button();
            _btnClose.Text = "關閉"; _btnClose.AutoSize = true; _btnClose.Padding = new Padding(16, 8, 16, 8);
            _btnClose.Anchor = AnchorStyles.Left; _btnClose.Margin = new Padding(12, 0, 0, 0);
            UiTheme.StyleButton(_btnClose);
            _btnClose.Click += delegate { Close(); };
            _lblStatus = new Label();
            _lblStatus.AutoSize = true; _lblStatus.ForeColor = UiTheme.TextSecondary;
            _lblStatus.Anchor = AnchorStyles.Left; _lblStatus.Margin = new Padding(20, 10, 0, 0);
            buttons.Controls.Add(_btnAdd);
            buttons.Controls.Add(_btnRemove);
            buttons.Controls.Add(_btnClose);
            buttons.Controls.Add(_lblStatus);
            Panel footerLine = new Panel();
            footerLine.Dock = DockStyle.Top; footerLine.Height = 1; footerLine.BackColor = UiTheme.Border;
            footer.Controls.Add(buttons);
            footer.Controls.Add(footerLine);
            Controls.Add(footer);

            // ===== 說明列（Top）=====
            Panel hintRow = new Panel();
            hintRow.Dock = DockStyle.Top; hintRow.AutoSize = true; hintRow.BackColor = UiTheme.BgMain;
            hintRow.Padding = new Padding(18, 10, 18, 2);
            Label hint = new Label();
            hint.Text = "目前的文件範本。可新增新範本或移除不需要的範本（會同步調整 Excel 欄位）。";
            hint.AutoSize = true; hint.Dock = DockStyle.Top; UiTheme.StyleHint(hint);
            hintRow.Controls.Add(hint);
            Controls.Add(hintRow);

            // ===== 頁首（Top，最後加入）=====
            Controls.Add(UiTheme.Header("範本管理", null));

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
