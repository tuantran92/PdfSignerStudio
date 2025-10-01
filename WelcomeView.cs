using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace PdfSignerStudio
{
    public class WelcomeView : UserControl
    {
        public event EventHandler? OpenFileClicked;
        public event EventHandler? LoadJsonClicked;
        public event EventHandler<string[]>? FilesDropped;

        private Color _accentColor = Color.FromArgb(0x25, 0x63, 0xEB);
        [Browsable(true)]
        public Color AccentColor { get => _accentColor; set { _accentColor = value; Invalidate(); foreach (var c in _cards) c.Accent = value; } }

        private bool _darkMode;
        [DefaultValue(false)]
        public bool DarkMode { get => _darkMode; set { _darkMode = value; ApplyTheme(); Invalidate(); foreach (var c in _cards) c.DarkMode = value; _drop.DarkMode = value; } }

        private readonly Label _title = new() { Text = "Welcome", AutoSize = true };
        private readonly Label _subtitle = new() { Text = "Open a file or load a JSON project", AutoSize = true };
        //private readonly FlowLayoutPanel _flow = new() { FlowDirection = FlowDirection.LeftToRight, WrapContents = false };

        private readonly Card _openFile;
        private readonly Card _loadJson;

        private readonly DropZone _drop = new();

        private readonly List<Card> _cards = new();

        public WelcomeView()
        {
            DoubleBuffered = true;
            BackColor = Color.FromArgb(245, 246, 248);

            _title.Font = new Font("Segoe UI Light", 28f);
            _subtitle.Font = new Font("Segoe UI", 11f);
            _subtitle.ForeColor = Color.FromArgb(80, 80, 80);

            _openFile = new Card("Open File", "Pick a PDF or DOCX");
            _loadJson = new Card("Load JSON", "Restore a saved project");

            _cards.AddRange(new[] { _openFile, _loadJson });
            foreach (var c in _cards) { c.Accent = _accentColor; c.Click += (s, e) => OnCardClick((Card)s!); }

            // Bỏ FlowLayoutPanel và thêm các card trực tiếp
            Controls.Add(_openFile);
            Controls.Add(_loadJson);

            _drop.Instructions = "Drag & drop PDF/DOCX here";
            _drop.FilesDropped += (s, paths) => FilesDropped?.Invoke(this, paths);

            Controls.AddRange(new Control[] { _title, _subtitle, _drop });
            _drop.BringToFront(); // Đảm bảo dropzone nằm trên cùng

            SizeChanged += (s, e) => DoLayout();
            Load += (s, e) => { ApplyTheme(); DoLayout(); };
        }

        private void OnCardClick(Card c)
        {
            if (c == _openFile) OpenFileClicked?.Invoke(this, EventArgs.Empty);
            else if (c == _loadJson) LoadJsonClicked?.Invoke(this, EventArgs.Empty);
        }

        private void DoLayout()
        {
            int w = ClientSize.Width;
            int h = ClientSize.Height;
            int defaultLeftPadding = 28;

            // === 1. TÍNH TOÁN VÙNG NỘI DUNG CHÍNH ĐỂ CĂN GIỮA TOÀN BỘ ===
            int contentW = Math.Min(980, Math.Max(720, w - defaultLeftPadding * 2));
            int contentAreaX = Math.Max(defaultLeftPadding, (w - contentW) / 2);

            // === 2. ĐẶT VỊ TRÍ CHO TITLE và SUBTITLE THEO VÙNG CĂN GIỮA ===
            _title.Location = new Point(contentAreaX + (contentW - _title.Width) / 2, 50);
            _subtitle.Location = new Point(contentAreaX + (contentW - _subtitle.Width) / 2, _title.Bottom + 5);

            // === 3. ĐẶT VỊ TRÍ CHO CÁC KHỐI "OPEN FILE" VÀ "LOAD JSON" ===
            var openFileW = 350;
            var loadJsonW = 350;
            var cardSpacing = 16;
            var totalCardsWidth = openFileW + loadJsonW + cardSpacing;

            var startX = contentAreaX + (contentW - totalCardsWidth) / 2;
            var cardY = 150;

            _openFile.Size = new Size(openFileW, 200);
            _loadJson.Size = new Size(loadJsonW, 200);

            _openFile.Location = new Point(startX, cardY);
            _loadJson.Location = new Point(startX + openFileW + cardSpacing, cardY);

            // === 4. ĐẶT VỊ TRÍ CHO DROP ZONE ===
            int dropW = contentW;
            int dropH = Math.Max(220, (int)((h - (_openFile.Bottom + 24) - 10) * 0.8));
            _drop.Size = new Size(dropW, dropH);
            _drop.Location = new Point(contentAreaX, _openFile.Bottom + 24);
        }

        public void RefreshRecent() { }

        private void ApplyTheme()
        {
            var bg = _darkMode ? Color.FromArgb(32, 34, 37) : Color.FromArgb(245, 246, 248);
            var fg = _darkMode ? Color.White : Color.Black;
            var sub = _darkMode ? Color.FromArgb(185, 188, 192) : Color.FromArgb(80, 80, 80);
            BackColor = bg; ForeColor = fg;
            _title.ForeColor = fg; _subtitle.ForeColor = sub;
            foreach (var c in _cards) c.DarkMode = _darkMode;
            _drop.DarkMode = _darkMode;
        }

        public class Card : Panel
        {
            public bool DarkMode { get; set; }
            public Color Accent { get; set; } = Color.FromArgb(0x25, 0x63, 0xEB);
            private bool _hover;
            private readonly Label _t = new(); private readonly Label _s = new();

            public Card(string title, string subtitle)
            {
                DoubleBuffered = true; Cursor = Cursors.Hand; Margin = new Padding(0, 0, 16, 0);
                _t.Text = title; _t.Font = new Font("Segoe UI Semibold", 13.5f); _t.AutoSize = true; _t.Location = new Point(25, 60);
                _s.Text = subtitle; _s.Font = new Font("Segoe UI", 10.5f); _s.AutoSize = true; _s.Location = new Point(25, 90); _s.ForeColor = Color.FromArgb(120, 124, 130);
                Controls.Add(_t); Controls.Add(_s);
                _t.Click += (s, e) => this.OnClick(e);
                _s.Click += (s, e) => this.OnClick(e);
                Paint += (s, e) => Draw(e.Graphics);
                MouseEnter += (s, e) => { _hover = true; Invalidate(); };
                MouseLeave += (s, e) => { _hover = false; Invalidate(); };
            }

            private void Draw(Graphics g)
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                var rect = ClientRectangle; rect.Inflate(-1, -1);
                float r = 18f;
                using GraphicsPath gp = Rounded(rect, r);
                using var br = new SolidBrush(DarkMode ? Color.FromArgb(_hover ? 48 : 42, 44, 48) : Color.White);
                using var pen = new Pen(DarkMode ? Color.FromArgb(58, 60, 64) : Color.FromArgb(230, 232, 236));
                g.FillPath(br, gp); g.DrawPath(pen, gp);
                if (_hover) { using var hl = new Pen(Accent, 2f); var inner = ClientRectangle; inner.Inflate(-2, -2); using var gp2 = Rounded(inner, r - 2); g.DrawPath(hl, gp2); }
            }
        }

        public class DropZone : Panel
        {
            public string Instructions { get; set; } = "Drag & drop files here";
            public bool DarkMode { get; set; }
            public event EventHandler<string[]>? FilesDropped;

            private readonly Image _dropIcon = Properties.Resources.move;

            public DropZone()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint |
                         ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.UserPaint,
                         true);
                DoubleBuffered = true;

                AllowDrop = true;
                Paint += (s, e) => Draw(e.Graphics);
                DragEnter += (s, e) => e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
                DragDrop += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) { string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop)!; FilesDropped?.Invoke(this, paths); } };
            }

            private void Draw(Graphics g)
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                var rect = ClientRectangle; rect.Inflate(-2, -2);
                using GraphicsPath gp = Rounded(rect, 16f);
                using var br = new SolidBrush(DarkMode ? Color.FromArgb(38, 40, 44) : Color.FromArgb(252, 253, 255));
                using var pen = new Pen(DarkMode ? Color.FromArgb(120, 160, 255) : Color.FromArgb(180, 186, 200), 2f) { DashStyle = DashStyle.Dash };
                g.FillPath(br, gp); g.DrawPath(pen, gp);

                using var font = new Font("Segoe UI Semibold", 10.5f);
                string text = Instructions;
                var textSize = g.MeasureString(text, font);

                var iconSize = new Size(32, 32);
                var totalHeight = iconSize.Height + 8 + textSize.Height;
                var startY = (Height - totalHeight) / 2f;

                var iconRect = new RectangleF((Width - iconSize.Width) / 2f, startY, iconSize.Width, iconSize.Height);
                g.DrawImage(_dropIcon, iconRect);

                var textRect = new RectangleF((Width - textSize.Width) / 2f, iconRect.Bottom + 8, textSize.Width, textSize.Height);
                using var textBrush = new SolidBrush(DarkMode ? Color.FromArgb(190, 195, 200) : Color.FromArgb(120, 124, 130));
                g.DrawString(text, font, textBrush, textRect);
            }
        }
        private static GraphicsPath Rounded(Rectangle r, float radius)
        {
            float d = radius * 2f; var gp = new GraphicsPath(); gp.StartFigure();
            gp.AddArc(r.X, r.Y, d, d, 180, 90); gp.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            gp.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90); gp.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            gp.CloseFigure(); return gp;
        }
    }
}