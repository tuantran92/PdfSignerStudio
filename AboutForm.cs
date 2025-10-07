using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace PdfSignerStudio
{
    public class AboutForm : Form
    {
        // Điền thông tin của bạn ở đây
        private const string CompanyNameText = "<Tên công ty>";
        private const string RepoUrl = "<URL repo source (public/private)>";

        private readonly string _baseDir = AppContext.BaseDirectory;

        private Panel leftPanel;
        private Panel rightPanel;

        private Label lblTitle, lblVersion, lblCompany, blurbLbl, lblViewerTitle;
        private LinkLabel lnkRepo, lnkReadme, lnkLicense, lnkThirdParty, lnkNotice;
        private RichTextBox viewer;
        private Button btnOpenExternal, btnCopy, btnClose;

        private string? _currentFilePath;

        public AboutForm()
        {
            // === Form cố định kích thước ===
            Text = "About PdfSignerStudio";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog; // ⬅️ không cho resize
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Font = new Font("Segoe UI", 9f);
            BackColor = Color.White;

            // Kích thước form cố định (đủ rộng để không che chữ)
            ClientSize = new Size(940, 560);

            // === Panel trái (thông tin + link) – kích thước cứng ===
            leftPanel = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(520, ClientSize.Height - 50), // chừa chỗ cho nút Close bên dưới (panel phải)
                BackColor = Color.White
            };
            Controls.Add(leftPanel);

            // === Panel phải (viewer) – kích thước cứng ===
            rightPanel = new Panel
            {
                Location = new Point(520, 0),
                Size = new Size(ClientSize.Width - 520, ClientSize.Height - 50),
                BackColor = Color.White
            };
            Controls.Add(rightPanel);

            // ---- LEFT CONTENT ----
            lblTitle = new Label
            {
                Text = "PdfSignerStudio",
                Font = new Font("Segoe UI Semibold", 17f),
                AutoSize = true,
                Location = new Point(18, 14)
            };
            leftPanel.Controls.Add(lblTitle);

            lblVersion = new Label
            {
                Text = $"Version {GetAppVersion()}",
                AutoSize = true,
                Location = new Point(22, 52)
            };
            leftPanel.Controls.Add(lblVersion);

            lblCompany = new Label
            {
                Text = CompanyNameText,
                AutoSize = true,
                ForeColor = Color.FromArgb(100, 100, 100),
                Location = new Point(22, 72)
            };
            leftPanel.Controls.Add(lblCompany);

            blurbLbl = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(480, 0), // giới hạn chiều ngang để tự xuống dòng
                Location = new Point(22, 106),
                Text =
@"Ứng dụng desktop nội bộ để mở DOCX/PDF, quét thẻ ký từ DOCX và đặt trường chữ ký lên PDF. 
Giao diện xem/chỉnh PDF hiển thị qua WebView2; phần xem PDF sử dụng pdf.js.

Ứng dụng liên kết iText 7 (AGPLv3). Xem các tệp giấy phép kèm theo bên dưới."
            };
            leftPanel.Controls.Add(blurbLbl);

            int linkX = 22; int linkY = blurbLbl.Bottom + 18; int step = 28;

            lnkRepo = CreateLink("Source code / Build guide (repo)", () => OpenUrl(RepoUrl), linkX, linkY); linkY += step;
            lnkReadme = CreateLink("README-AGPL.md", () => LoadDocument("README-AGPL.md"), linkX, linkY); linkY += step;
            lnkLicense = CreateLink("LICENSE-AGPL (AGPLv3)", () => LoadDocument("LICENSE-AGPL"), linkX, linkY); linkY += step;
            lnkThirdParty = CreateLink("THIRD-PARTY-NOTICES.md", () => LoadDocument("THIRD-PARTY-NOTICES.md"), linkX, linkY); linkY += step;
            lnkNotice = CreateLink("NOTICE.txt", () => LoadDocument("NOTICE.txt"), linkX, linkY);

            leftPanel.Controls.AddRange(new Control[] { lnkRepo, lnkReadme, lnkLicense, lnkThirdParty, lnkNotice });

            // ---- RIGHT CONTENT ----
            lblViewerTitle = new Label
            {
                Text = "Viewer",
                AutoSize = true,
                Font = new Font("Segoe UI Semibold", 10.5f),
                Location = new Point(12, 10)
            };
            rightPanel.Controls.Add(lblViewerTitle);

            viewer = new RichTextBox
            {
                Location = new Point(12, 36),
                Size = new Size(rightPanel.Width - 24, rightPanel.Height - 90), // cố định theo panel phải
                ReadOnly = true,
                BorderStyle = BorderStyle.FixedSingle,
                DetectUrls = true,
                WordWrap = true,
                Font = new Font("Consolas", 10f),
                ScrollBars = RichTextBoxScrollBars.Both
            };
            rightPanel.Controls.Add(viewer);

            btnOpenExternal = new Button
            {
                Text = "Open file",
                Size = new Size(96, 30),
                Location = new Point(12, rightPanel.Height - 44),
            };
            btnOpenExternal.Click += (_, __) =>
            {
                if (_currentFilePath != null)
                {
                    try { Process.Start(new ProcessStartInfo(_currentFilePath) { UseShellExecute = true }); }
                    catch { MessageBox.Show("Không mở được file bằng ứng dụng ngoài.", "Open file", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                }
            };
            rightPanel.Controls.Add(btnOpenExternal);

            btnCopy = new Button
            {
                Text = "Copy",
                Size = new Size(96, 30),
                Location = new Point(116, rightPanel.Height - 44),
            };
            btnCopy.Click += (_, __) =>
            {
                if (!string.IsNullOrEmpty(viewer.Text))
                    Clipboard.SetText(viewer.Text);
            };
            rightPanel.Controls.Add(btnCopy);

            // ---- CLOSE BUTTON (cố định góc phải dưới của Form) ----
            btnClose = new Button
            {
                Text = "Close",
                DialogResult = DialogResult.OK,
                Size = new Size(96, 30),
                Location = new Point(ClientSize.Width - 96 - 12, ClientSize.Height - 36 - 12)
            };
            Controls.Add(btnClose);
            AcceptButton = btnClose;

            // Load mặc định
            Load += (_, __) =>
            {
                LoadDocument("README-AGPL.md");
            };
        }

        // ===== Helpers =====
        private static LinkLabel CreateLink(string text, Action onClick, int x, int y)
        {
            var lnk = new LinkLabel
            {
                Text = text,
                AutoSize = true,
                Location = new Point(x, y),
                LinkBehavior = LinkBehavior.HoverUnderline
            };
            lnk.LinkColor = Color.FromArgb(0x25, 0x63, 0xEB);
            lnk.ActiveLinkColor = Color.FromArgb(0x1D, 0x4E, 0xBA);
            lnk.Click += (_, __) => onClick();
            return lnk;
        }

        private static string GetAppVersion()
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var fvi = FileVersionInfo.GetVersionInfo(asm.Location);
                if (!string.IsNullOrWhiteSpace(fvi.FileVersion)) return fvi.FileVersion;
            }
            catch { }
            return Application.ProductVersion;
        }

        private void OpenUrl(string? url)
        {
            if (string.IsNullOrWhiteSpace(url)) return;
            try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
            catch { MessageBox.Show("Không mở được URL.", "Open URL", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }

        private void LoadDocument(string relativeFile)
        {
            string? path = null;
            try
            {
                path = Path.Combine(_baseDir, relativeFile);
                if (!File.Exists(path))
                {
                    var fallback = Directory.GetFiles(_baseDir, relativeFile, SearchOption.AllDirectories);
                    if (fallback.Length > 0) path = fallback[0];
                }

                if (path == null || !File.Exists(path))
                {
                    viewer.Clear();
                    lblViewerTitle.Text = $"{relativeFile} (not found)";
                    _currentFilePath = null;
                    return;
                }

                var text = File.ReadAllText(path, new UTF8Encoding(false));
                viewer.Clear();
                viewer.Text = text;
                lblViewerTitle.Text = Path.GetFileName(path);
                _currentFilePath = path;
            }
            catch
            {
                viewer.Clear();
                lblViewerTitle.Text = $"{relativeFile} (error)";
                _currentFilePath = null;
            }
        }
    }
}
