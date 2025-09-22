using System.Drawing;
using System.Windows.Forms;

namespace PdfSignerStudio
{
    public partial class SplashForm : Form
    {
        public SplashForm()
        {
            InitializeComponent(); // Hàm này được designer sinh ra, cứ để đó

            // --- Tùy chỉnh giao diện cho Splash Form ---

            // Thiết lập cho Form
            this.FormBorderStyle = FormBorderStyle.None; // Không có viền, nút đóng/mở
            this.StartPosition = FormStartPosition.CenterScreen; // Luôn hiện ở giữa màn hình
            this.ClientSize = new Size(400, 100);
            this.BackColor = Color.White;
            this.Padding = new Padding(10);
            this.TopMost = true; // Luôn hiển thị trên cùng

            // Tạo Label để hiển thị thông báo
            Label lblMessage = new Label
            {
                Text = "Converting DOCX to PDF...",
                Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 0),
                ForeColor = Color.FromArgb(64, 64, 64),
                Location = new Point(20, 20),
                AutoSize = true
            };

            // Tạo ProgressBar với hiệu ứng chạy qua lại
            ProgressBar progressBar = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30,
                Location = new Point(20, 55),
                Width = 360,
                Height = 15
            };

            // Thêm các control vào Form
            this.Controls.Add(lblMessage);
            this.Controls.Add(progressBar);
        }
    }
}