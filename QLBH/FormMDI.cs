using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QLBH
{
    public partial class FormMDI : Form
    {
        public FormMDI()
        {
            InitializeComponent();
        }

        private void tàiKhoảnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DongTatCaFormCon();
            FormTaiKhoan F = new FormTaiKhoan();
            F.MdiParent = this;
            F.Show();
            
        }

        private void nhânViênToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DongTatCaFormCon();
            FormNhanVien F = new FormNhanVien();
            F.MdiParent = this;
            F.Show();
           
        }

        private void đăngKýTKToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DongTatCaFormCon();
            FormDangKy F = new FormDangKy();
            F.MdiParent = this;
            F.Show();
        }

        private void hóaĐơnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DongTatCaFormCon();
            FormHD F = new FormHD();
            F.MdiParent = this;
            F.Show();
        }
        private void DongTatCaFormCon()
        {
            foreach (Form childForm in this.MdiChildren)
            {
                childForm.Close();
            }
        }

        private void kháchHàngToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DongTatCaFormCon();
            FormKhachHang F = new FormKhachHang();
            F.MdiParent = this;
            F.Show();
        }

        private void tàiKhoảnToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            DongTatCaFormCon();
            FormTaiKhoan1 F = new FormTaiKhoan1();
            F.MdiParent = this;
            F.Show();
        }

        private void thôngTInToolStripMenuItem_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void chiTiếtHDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DongTatCaFormCon();
            FormCTHD F = new FormCTHD();
            F.MdiParent = this;
            F.Show();
        }

        private void FormMDI_Load(object sender, EventArgs e)
        {
            
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;   
            pictureBox1.Dock = DockStyle.Fill;
            pictureBox1.Visible = true;

            this.MdiChildActivate += new EventHandler(FormMDI_MdiChildActivate);
        }
        private void FormMDI_MdiChildActivate(object sender, EventArgs e)
        {
            if (this.ActiveMdiChild != null)
            {
                pictureBox1.Visible = false;
            }
            else
            {
               
                pictureBox1.Visible = true;
            }
        }

        private void hàngHóaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DongTatCaFormCon();
            FormHangHoa F = new FormHangHoa();
            F.MdiParent = this;
            F.Show();
        }

        private void chấtLiệuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DongTatCaFormCon();
            FormChatLieu F = new FormChatLieu();
            F.MdiParent = this;
            F.Show();
        }
    }
}
