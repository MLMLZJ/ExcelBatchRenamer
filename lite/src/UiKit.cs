using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelBatchRename
{
    /// <summary>统一的控件构造工具，保证两种 DPI 下排版一致</summary>
    internal static class UiKit
    {
        public static readonly Label[] Noop = new Label[0];
        public static readonly Font UiFont = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
        public static readonly Font LogFont = new Font("Consolas", 10F, FontStyle.Regular, GraphicsUnit.Point, 134);
        public static readonly Color HintColor = Color.FromArgb(0x60, 0x60, 0x60);

        private static Icon _appIcon;

        /// <summary>取嵌入的程序图标（不再打开自身 exe 读资源）</summary>
        public static Icon AppIcon()
        {
            if (_appIcon != null) return _appIcon;
            try
            {
                System.Reflection.Assembly asm = System.Reflection.Assembly.GetExecutingAssembly();
                using (System.IO.Stream s = asm.GetManifestResourceStream("app.ico"))
                {
                    if (s != null) _appIcon = new Icon(s);
                }
            }
            catch
            {
            }
            if (_appIcon == null)
            {
                try { _appIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
                catch { }
            }
            return _appIcon;
        }

        public static Label Text(string text, int width)
        {
            Label l = new Label();
            l.Text = text;
            l.AutoSize = true;
            l.TextAlign = ContentAlignment.MiddleLeft;
            l.Margin = new Padding(0, 4, 6, 4);
            return l;
        }

        public static Label Hint(string text)
        {
            Label l = Text(text, 0);
            l.ForeColor = HintColor;
            return l;
        }

        /// <summary>输入框。width &lt;= 0 表示不固定宽度，交给布局容器拉伸</summary>
        public static TextBox Input(int width)
        {
            TextBox t = new TextBox();
            if (width > 0) t.Width = width;
            t.Margin = new Padding(0, 4, 6, 4);
            return t;
        }

        public static Button Button(string text, EventHandler onClick)
        {
            Button b = new Button();
            b.Text = text;
            b.AutoSize = true;
            b.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            b.Padding = new Padding(10, 4, 10, 4);
            b.Margin = new Padding(0, 4, 8, 4);
            b.UseVisualStyleBackColor = true;
            b.Click += onClick;
            return b;
        }

        /// <summary>一行控件（自动宽度、不换行）</summary>
        public static FlowLayoutPanel Row(params Control[] items)
        {
            FlowLayoutPanel p = new FlowLayoutPanel();
            p.AutoSize = true;
            p.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            p.WrapContents = false;
            p.Margin = new Padding(0);
            p.Padding = new Padding(0);
            p.Controls.AddRange(items);
            return p;
        }

        /// <summary>竖排容器（自动宽度）</summary>
        public static FlowLayoutPanel Column(params Control[] items)
        {
            FlowLayoutPanel p = new FlowLayoutPanel();
            p.AutoSize = true;
            p.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            p.FlowDirection = FlowDirection.TopDown;
            p.WrapContents = false;
            p.Margin = new Padding(0);
            p.Padding = new Padding(0);
            p.Controls.AddRange(items);
            return p;
        }

        public static RadioButton Radio(string text, bool isChecked, EventHandler onChecked)
        {
            RadioButton r = new RadioButton();
            r.Text = text;
            r.AutoSize = true;
            r.Checked = isChecked;
            r.Margin = new Padding(0, 4, 16, 4);
            if (onChecked != null) r.CheckedChanged += onChecked;
            return r;
        }

        public static CheckBox Check(string text, bool isChecked)
        {
            CheckBox c = new CheckBox();
            c.Text = text;
            c.AutoSize = true;
            c.Checked = isChecked;
            c.Margin = new Padding(0, 6, 8, 4);
            return c;
        }

        /// <summary>
        /// 带标题的分区块。用 TableLayoutPanel 保证内容能随窗口横向拉伸
        /// （FlowLayoutPanel 是自动宽的，子控件 Dock=Fill 撑不开它）。
        /// </summary>
        public static TableLayoutPanel Section(string title, Control content)
        {
            Label head = new Label();
            head.Text = title;
            head.AutoSize = true;
            head.Font = new Font(UiFont.FontFamily, 9.5F, FontStyle.Bold, GraphicsUnit.Point, 134);
            head.ForeColor = Color.FromArgb(0x1F, 0x3B, 0x63);
            head.Margin = new Padding(0, 8, 0, 3);

            Panel line = new Panel();
            line.Height = 1;
            line.BackColor = Color.FromArgb(0xD0, 0xD7, 0xE2);
            line.Dock = DockStyle.Top;
            line.Margin = new Padding(0, 0, 0, 4);

            TableLayoutPanel box = new TableLayoutPanel();
            box.Dock = DockStyle.Fill;
            box.AutoSize = true;
            box.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            box.ColumnCount = 1;
            box.RowCount = 3;
            box.Margin = new Padding(0);
            box.Padding = new Padding(0);
            box.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            box.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            box.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            box.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            content.Dock = DockStyle.Fill;
            content.Margin = new Padding(0);

            box.Controls.Add(head, 0, 0);
            box.Controls.Add(line, 0, 1);
            box.Controls.Add(content, 0, 2);
            return box;
        }

        public static GroupBox Group(string title, Control content)
        {
            GroupBox g = new GroupBox();
            g.Text = title;
            g.AutoSize = true;
            g.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            g.Dock = DockStyle.Fill;
            g.Padding = new Padding(10, 6, 10, 10);
            g.Margin = new Padding(0, 4, 0, 4);
            g.Controls.Add(content);
            return g;
        }


        /// <summary>
        /// 响应式表单：标签列固定宽，内容列按百分比拉伸。
        /// 关键：按钮、提示都放在"内容列"里，作为一个横向 FlowLayoutPanel，
        /// 这样它们不会和百分比列抢空间而互相挤压。
        /// </summary>
        public static TableLayoutPanel Form(int labelWidth)
        {
            TableLayoutPanel t = new TableLayoutPanel();
            t.Dock = DockStyle.Fill;
            t.AutoSize = true;
            t.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            t.ColumnCount = 2;
            t.Margin = new Padding(0);
            t.Padding = new Padding(0);
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            return t;
        }

        /// <summary>
        /// 往表单加一行：标签 + 内容。
        /// content 里可以随便塞输入框/按钮/提示，用 Row() 组织即可；
        /// 需要输入框拉伸时用 <see cref="Stretch"/> 包一层。
        /// </summary>
        public static void AddField(TableLayoutPanel form, int row, string label, Control content)
        {
            form.RowCount = row + 1;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            if (label != null)
            {
                Label l = Text(label, 0);
                l.AutoSize = true;
                l.TextAlign = ContentAlignment.MiddleLeft;
                l.Margin = new Padding(0, 4, 12, 4);
                form.Controls.Add(l, 0, row);
            }

            if (content != null)
            {
                content.Dock = DockStyle.Fill;
                content.Margin = new Padding(0, 3, 0, 3);
                form.Controls.Add(content, 1, row);
            }
        }

        /// <summary>占满整行的控件</summary>
        public static void AddFullRow(TableLayoutPanel form, int row, Control control)
        {
            form.RowCount = row + 1;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            control.Anchor = AnchorStyles.Left;
            control.Margin = new Padding(0, 6, 8, 4);
            form.Controls.Add(control, 1, row);
        }

        /// <summary>让输入框在可用宽度内拉伸（配合 AddField 使用）</summary>
        public static TableLayoutPanel Stretch(Control input, int minWidth = 0)
        {
            TableLayoutPanel t = new TableLayoutPanel();
            t.Dock = DockStyle.Fill;
            t.ColumnCount = 1;
            t.RowCount = 1;
            t.Margin = new Padding(0);
            t.Padding = new Padding(0);
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            t.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            input.Dock = DockStyle.Fill;
            input.Margin = new Padding(0);
            if (minWidth > 0) input.MinimumSize = new Size(minWidth, 26);
            else input.MinimumSize = new Size(0, 26);
            t.Controls.Add(input, 0, 0);
            return t;
        }

        /// <summary>横向排列：输入框(拉伸) + 按钮 + 提示</summary>
        public static TableLayoutPanel Field(Control input, Control button = null, Control hint = null,
                                            int hintWidth = 0)
        {
            TableLayoutPanel t = new TableLayoutPanel();
            t.Dock = DockStyle.Fill;
            t.AutoSize = true;
            t.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            t.ColumnCount = 3;
            t.RowCount = 1;
            t.Margin = new Padding(0);
            t.Padding = new Padding(0);
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(hintWidth > 0 ? SizeType.Absolute : SizeType.AutoSize,
                                               hintWidth > 0 ? hintWidth : 0F));
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            if (input != null)
            {
                input.Dock = DockStyle.Fill;
                input.MinimumSize = new Size(0, 26);
                input.Margin = new Padding(0, 0, 8, 0);
                t.Controls.Add(input, 0, 0);
            }
            if (button != null)
            {
                button.AutoSize = true;
                button.Margin = new Padding(0, 0, 8, 0);
                t.Controls.Add(button, 1, 0);
            }
            if (hint != null)
            {
                hint.Anchor = AnchorStyles.Left;
                hint.Margin = new Padding(0);
                t.Controls.Add(hint, 2, 0);
            }
            return t;
        }

        /// <summary>
        /// 给内容区加最大宽度限制：窗口超宽时左右留白，内容居中。
        /// 否则 1500px 宽的窗口会出现横跨全屏的输入框，很难看。
        /// </summary>
        public static Control MaxWidth(Control inner, int maxWidth)
        {
            TableLayoutPanel t = new TableLayoutPanel();
            t.Dock = DockStyle.Fill;
            t.ColumnCount = 3;
            t.RowCount = 1;
            t.Margin = new Padding(0);
            t.Padding = new Padding(0);
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, maxWidth));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            inner.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            t.Controls.Add(inner, 1, 0);
            return t;
        }

        public static TableLayoutPanel Root()
        {
            TableLayoutPanel t = new TableLayoutPanel();
            t.Dock = DockStyle.Fill;
            t.Padding = new Padding(12, 10, 12, 10);
            t.ColumnCount = 1;
            t.RowCount = 4;
            t.AutoScroll = false;
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            t.Margin = new Padding(0);
            t.Padding = new Padding(12, 10, 12, 10);
            return t;
        }
    }
}