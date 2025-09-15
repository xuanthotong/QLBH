using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace QLBH
{
    public partial class FormTaiKhoan : Form
    {
        public FormTaiKhoan()
        {
            InitializeComponent();
        }
        string Nguon = @"Data Source=DESKTOP-FFTC51V\SQLEXPRESS;Initial Catalog=QuanLyBH;Integrated Security=True";
        string Lenh = @"";
        SqlConnection KetNoi;
        SqlCommand ThucHien;
        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBoxTK.Text))
            {
                MessageBox.Show("Vui lòng nhập tên đăng nhập!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBoxTK.Focus();
                return;
            }
            if (string.IsNullOrWhiteSpace(textBoxMK.Text))
            {
                MessageBox.Show("Vui lòng nhập mật khẩu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBoxMK.Focus();
                return; 
            }

            KetNoi = new SqlConnection(Nguon);
            Lenh = @"SELECT COUNT(*) FROM TaiKhoan 
             WHERE TenDangNhap = @TenDangNhap AND MatKhau = @MatKhau";
            KetNoi.Open();
            ThucHien = new SqlCommand(Lenh, KetNoi);
            ThucHien.Parameters.Add("@TenDangNhap", SqlDbType.NVarChar);
            ThucHien.Parameters["@TenDangNhap"].Value = textBoxTK.Text;
            ThucHien.Parameters.Add("@MatKhau", SqlDbType.NVarChar);
            ThucHien.Parameters["@MatKhau"].Value = textBoxMK.Text;
            int count = (int) ThucHien.ExecuteScalar(); 
            KetNoi.Close();
            if (count > 0)
            {
                MessageBox.Show("Đăng nhập thành công!");
                this.Hide();
                FormMDI F = new FormMDI(); 
                F.Show(); 
            }
            else
            {
                MessageBox.Show("Tên đăng nhập hoặc mật khẩu không đúng.");
            }
        }

        private void buttonThoat_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn thoát hệ thống?",
                                        "Xác nhận",
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormDangKy F = new FormDangKy();
            F.Show();
            this.Hide();       
        }

        private void checkBoxMK_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxMK.Checked)
            {
                textBoxMK.UseSystemPasswordChar = false;
            }
            else
            {
                textBoxMK.UseSystemPasswordChar = true;
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormQMK F = new FormQMK();
            F.Show();
            this.Hide(); 
        }
    }
}
