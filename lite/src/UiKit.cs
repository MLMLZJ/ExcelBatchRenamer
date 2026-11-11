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
        public static readonly Font LogFont = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
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

        public static TextBox Input(int width)
        {
            TextBox t = new TextBox();
            t.Width = width;
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

        /// <summary>带标题的分区块（用加粗标签代替 GroupBox，避免高 DPI 下标题被裁切）</summary>
        public static FlowLayoutPanel Section(string title, Control content)
        {
            Label head = new Label();
            head.Text = title;
            head.AutoSize = true;
            head.Font = new Font(UiFont.FontFamily, 9.5F, FontStyle.Bold, GraphicsUnit.Point, 134);
            head.ForeColor = Color.FromArgb(0x1F, 0x3B, 0x63);
            head.Margin = new Padding(0, 8, 0, 2);

            Panel line = new Panel();
            line.Height = 1;
            line.BackColor = Color.FromArgb(0xD0, 0xD7, 0xE2);
            line.Width = 820;
            line.Margin = new Padding(0, 0, 0, 2);

            FlowLayoutPanel p = new FlowLayoutPanel();
            p.AutoSize = true;
            p.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            p.FlowDirection = FlowDirection.TopDown;
            p.WrapContents = false;
            p.Margin = new Padding(0);
            p.Padding = new Padding(0);
            p.Controls.Add(head);
            p.Controls.Add(line);
            p.Controls.Add(content);
            return p;
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