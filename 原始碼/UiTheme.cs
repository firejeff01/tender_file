using System;
using System.Drawing;
using System.Windows.Forms;

namespace TenderDocGen
{
    /// <summary>
    /// 「Warm Professional」暖色主題（沿用 Tender.Desktop 專案色票），
    /// 以 WinForms 近似 WPF 的視覺：暖白背景、焦糖陶土強調色、圓潤扁平按鈕、卡片感表格。
    /// </summary>
    static class UiTheme
    {
        // ---- 背景層 ----
        public static readonly Color BgMain = Hex("#FBF7EF");
        public static readonly Color BgCard = Hex("#FFFCF7");
        public static readonly Color BgAccent = Hex("#F0E6D2");
        public static readonly Color BgSubtle = Hex("#F5EFE3");
        // ---- 邊線 ----
        public static readonly Color Border = Hex("#E0D4BC");
        public static readonly Color BorderStrong = Hex("#C9B89C");
        // ---- 主強調色（焦糖陶土）----
        public static readonly Color Primary = Hex("#C4823C");
        public static readonly Color PrimaryDark = Hex("#9B5C24");
        public static readonly Color PrimaryLight = Hex("#F4E6CD");
        public static readonly Color PrimaryHover = Hex("#FBF1DE");
        // ---- 文字 ----
        public static readonly Color TextPrimary = Hex("#3A2E20");
        public static readonly Color TextSecondary = Hex("#7A6857");
        public static readonly Color TextMuted = Hex("#A89888");
        public static readonly Color TextOnPrimary = Hex("#FFFCF7");
        // ---- 語意色 ----
        public static readonly Color Success = Hex("#7A9963");
        public static readonly Color SuccessDark = Hex("#5C7A48");
        public static readonly Color Warning = Hex("#D69830");
        public static readonly Color Error = Hex("#B5523B");
        public static readonly Color Info = Hex("#5A7290");

        public const string FontName = "Microsoft JhengHei UI";

        public static Font Font(float size) { return new Font(FontName, size); }
        public static Font FontBold(float size) { return new Font(FontName, size, FontStyle.Bold); }

        static Color Hex(string h) { return ColorTranslator.FromHtml(h); }

        // ==================== 表單 ====================
        public static void StyleForm(Form f)
        {
            f.BackColor = BgMain;
            f.ForeColor = TextPrimary;
            f.Font = Font(12f);
        }

        // ==================== 按鈕 ====================
        // 次要（預設）按鈕：卡片底、暖邊、hover 變主色淡底
        public static void StyleButton(Button b)
        {
            b.FlatStyle = FlatStyle.Flat;
            b.BackColor = BgCard;
            b.ForeColor = TextPrimary;
            b.FlatAppearance.BorderColor = BorderStrong;
            b.FlatAppearance.BorderSize = 1;
            b.FlatAppearance.MouseOverBackColor = PrimaryHover;
            b.FlatAppearance.MouseDownBackColor = PrimaryLight;
            b.Font = Font(12f);
            b.ForeColor = TextPrimary;
            b.Cursor = Cursors.Hand;
            b.UseVisualStyleBackColor = false;
            if (b.Padding == Padding.Empty) b.Padding = new Padding(14, 6, 14, 6);
        }

        static void StyleColored(Button b, Color bg, Color border, Color hover)
        {
            StyleButton(b);
            b.BackColor = bg;
            b.ForeColor = TextOnPrimary;
            b.Font = FontBold(12f);
            b.FlatAppearance.BorderColor = border;
            b.FlatAppearance.MouseOverBackColor = hover;
            b.FlatAppearance.MouseDownBackColor = border;
        }

        public static void StylePrimary(Button b) { StyleColored(b, Primary, PrimaryDark, PrimaryDark); }
        public static void StyleSuccess(Button b) { StyleColored(b, Success, SuccessDark, SuccessDark); }
        public static void StyleDanger(Button b) { StyleColored(b, Error, Hex("#8E3E2D"), Hex("#8E3E2D")); }
        public static void StyleInfo(Button b) { StyleColored(b, Info, Hex("#42566E"), Hex("#42566E")); }

        // ==================== 文字方塊 ====================
        public static void StyleTextBox(TextBox t)
        {
            t.BorderStyle = BorderStyle.FixedSingle;
            t.BackColor = BgCard;
            t.ForeColor = TextPrimary;
            t.Font = Font(12f);
        }

        // ==================== 標籤 ====================
        public static Label Title(string text)
        {
            Label l = new Label();
            l.Text = text;
            l.AutoSize = true;
            l.Font = FontBold(16.5f);
            l.ForeColor = TextPrimary;
            return l;
        }

        public static void StyleHint(Label l)
        {
            l.ForeColor = TextSecondary;
            l.Font = Font(11.5f);
        }

        // ==================== 頁首（accent bar + 標題 + 工具列）====================
        // 回傳 Dock=Top、AutoSize（隨字型/DPI 縮放，不用絕對高度）的卡片色頁首列。
        // title 靠左（前有陶土色直條），rightPanel 靠右（工具列），可傳 null。
        public static Control Header(string title, Control rightPanel)
        {
            Panel wrap = new Panel();
            wrap.Dock = DockStyle.Top;
            wrap.AutoSize = true;
            wrap.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            wrap.BackColor = BgCard;

            TableLayoutPanel bar = new TableLayoutPanel();
            bar.Dock = DockStyle.Top;
            bar.AutoSize = true;
            bar.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            bar.BackColor = BgCard;
            bar.ColumnCount = 4;
            bar.RowCount = 1;
            bar.Padding = new Padding(16, 12, 16, 12);
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            Panel accent = new Panel();
            accent.Size = new Size(5, 30);
            accent.BackColor = Primary;
            accent.Anchor = AnchorStyles.None;
            accent.Margin = new Padding(0, 2, 12, 2);
            Label t = Title(title);
            t.Anchor = AnchorStyles.Left;
            t.Margin = new Padding(0, 2, 0, 2);
            bar.Controls.Add(accent, 0, 0);
            bar.Controls.Add(t, 1, 0);
            if (rightPanel != null)
            {
                rightPanel.Anchor = AnchorStyles.Right;
                rightPanel.BackColor = Color.Transparent;
                bar.Controls.Add(rightPanel, 3, 0);
            }

            Panel bottomLine = new Panel();
            bottomLine.Dock = DockStyle.Bottom;
            bottomLine.Height = 1;
            bottomLine.BackColor = Border;

            wrap.Controls.Add(bar);
            wrap.Controls.Add(bottomLine);
            return wrap;
        }

        // ==================== 卡片面板 ====================
        public static Panel Card()
        {
            Panel p = new Panel();
            p.BackColor = BgCard;
            p.Padding = new Padding(12);
            return p;
        }

        // ==================== 表格 ====================
        public static void StyleGrid(DataGridView g)
        {
            g.EnableHeadersVisualStyles = false;
            g.BackgroundColor = BgCard;
            g.BorderStyle = BorderStyle.None;
            g.GridColor = Border;
            g.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            g.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            g.RowHeadersVisible = false;
            g.AllowUserToResizeRows = false;
            g.Font = Font(12f);

            DataGridViewCellStyle head = g.ColumnHeadersDefaultCellStyle;
            head.BackColor = BgAccent;
            head.ForeColor = TextPrimary;
            head.Font = FontBold(12f);
            head.SelectionBackColor = BgAccent;
            head.SelectionForeColor = TextPrimary;
            head.Padding = new Padding(6, 8, 6, 8);
            g.ColumnHeadersHeight = 44;
            g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            DataGridViewCellStyle cell = g.DefaultCellStyle;
            cell.BackColor = BgCard;
            cell.ForeColor = TextPrimary;
            cell.SelectionBackColor = PrimaryLight;
            cell.SelectionForeColor = TextPrimary;
            cell.Padding = new Padding(4, 5, 4, 5);
            g.AlternatingRowsDefaultCellStyle.BackColor = BgSubtle;
            g.AlternatingRowsDefaultCellStyle.SelectionBackColor = PrimaryLight;
            g.AlternatingRowsDefaultCellStyle.SelectionForeColor = TextPrimary;
            g.RowTemplate.Height = 38;
        }
    }
}
