using System.Drawing;
using System.Windows.Forms;

namespace PdfSignerStudio
{
    public partial class DimmerForm : Form
    {
        public DimmerForm()
        {
            InitializeComponent();

            // Cấu hình để form này trở thành một lớp phủ mờ
            this.FormBorderStyle = FormBorderStyle.None;
            this.BackColor = Color.Black;
            this.Opacity = 0.6; // Độ mờ, 0.6 là 60%
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.Manual;
        }
    }
}