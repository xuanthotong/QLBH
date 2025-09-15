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
    public partial class FormDangKy : Form
    {
        public FormDangKy()
        {
            InitializeComponent();
        }
        string Nguon = @"Data Source=DESKTOP-FFTC51V\SQLEXPRESS;Initial Catalog=QuanLyBH;Integrated Security=True";
        string Lenh = @"";
        SqlConnection KetNoi;
        SqlCommand ThucHien;
        private void button1_Click(object sender, EventArgs e)
        {
            KetNoi = new SqlConnection(Nguon);
            if (string.IsNullOrWhiteSpace(textBoxTK.Text) || string.IsNullOrWhiteSpace(textBoxMK.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ tên đăng nhập và mật khẩu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Lenh = @"SELECT COUNT(*) 
                  FROM TaiKhoan
                  WHERE TenDangNhap = @TenDangNhap";
              KetNoi.Open();
            using (SqlCommand check = new SqlCommand(Lenh, KetNoi))
            {
                check.Parameters.Add("@TenDangNhap", SqlDbType.NVarChar).Value = textBoxTK.Text.Trim();
                int count0 = (int)check.ExecuteScalar();
                if (count0 > 0)
                {
                    MessageBox.Show("Tên đăng nhập đã tồn tại! Vui lòng chọn tên khác.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
           
            Lenh = @"INSERT INTO TaiKhoan
                  (TenDangNhap, MatKhau)
             VALUES (@TenDangNhap,@MatKhau)";
            ThucHien = new SqlCommand(Lenh, KetNoi);
            ThucHien.Parameters.Add("@TenDangNhap", SqlDbType.NVarChar);
            ThucHien.Parameters.Add("@MatKhau", SqlDbType.NVarChar);
            ThucHien.Parameters["@TenDangNhap"].Value = textBoxTK.Text;       
            ThucHien.Parameters["@Matkhau"].Value = textBoxMK.Text;
            int count = ThucHien.ExecuteNonQuery();
            KetNoi.Close();
            if (count > 0)
            {
                MessageBox.Show("Đăng ký thành công!");
                FormTaiKhoan F = new FormTaiKhoan();
                F.Show();
            }
            else
            {
                MessageBox.Show("Tên đăng ký không tồn tại");
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
                this.Hide();
                FormTaiKhoan F = new FormTaiKhoan();
                F.Show(); 
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBoxMK.UseSystemPasswordChar = false;
            }
            else
            {
               textBoxMK.UseSystemPasswordChar = true;
            }
        }
    }
}
