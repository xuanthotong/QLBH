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
    public partial class FormTaiKhoan1 : Form
    {
        public FormTaiKhoan1()
        {
            InitializeComponent();
        }
        string Nguon = @"Data Source=DESKTOP-FFTC51V\SQLEXPRESS;Initial Catalog=QuanLyBH;Integrated Security=True";
        string Lenh = @"";
        SqlConnection KetNoi;
        SqlCommand ThucHien;
        SqlDataReader Doc;
        private void button1_Click(object sender, EventArgs e)
        {

        }
        private string tenDangNhapHienTai;

        public FormTaiKhoan1(string tenDangNhap)
        {
            InitializeComponent();
            tenDangNhapHienTai = tenDangNhap; 
        }
        private void buttonXoa_Click(object sender, EventArgs e)
        {
            KetNoi = new SqlConnection(Nguon);
            DialogResult D = MessageBox.Show("Bạn có muốn xóa nhân viên này không" + comboBoxTaiKhoan.Text, " Chú ý", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (D == DialogResult.Yes)
            {
              
                Lenh = @"DELETE FROM TaiKhoan 
                      WHERE TenDangNhap = @TenDangNhap";
                ThucHien = new SqlCommand(Lenh, KetNoi);
                ThucHien.Parameters.Add("@TenDangNhap", SqlDbType.NVarChar).Value = dataGridView.CurrentRow.Cells[0].Value;
                KetNoi.Open();
                int result = ThucHien.ExecuteNonQuery();
                KetNoi.Close();
                if (result > 0)
                {
                    MessageBox.Show("Xóa tài khoản thành công! Bạn sẽ được đăng xuất.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);                 
                    this.Hide();
                    FormTaiKhoan formDangNhap = new FormTaiKhoan();
                    formDangNhap.ShowDialog();
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Không tìm thấy tài khoản để xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {

            }
        }
        void HienThi()
        {
            dataGridView.Rows.Clear();
            Lenh = @"SELECT TenDangNhap
                     FROM     TaiKhoan";
            ThucHien = new SqlCommand(Lenh, KetNoi);
            KetNoi.Open();
            Doc = ThucHien.ExecuteReader();
            int i = 0;
            while (Doc.Read())
            {
                dataGridView.Rows.Add();
                dataGridView.Rows[i].Cells[0].Value = Doc[0];
                i++;
            }
            KetNoi.Close();
        }

        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            comboBoxTaiKhoan.Text = dataGridView.CurrentRow.Cells[0].Value.ToString();
        }

        private void FormTaiKhoan1_Load(object sender, EventArgs e)
        {
            KetNoi = new SqlConnection(Nguon);
            HienThi();
            comboBoxTaiKhoan.Items.Clear();
            Lenh = @"SELECT TenDangNhap
                   FROM TaiKhoan";
            ThucHien = new SqlCommand(Lenh, KetNoi);
            KetNoi.Open();
            Doc = ThucHien.ExecuteReader();
            int i = 0;
            while (Doc.Read())
            {
                comboBoxTaiKhoan.Items.Add(Doc[0]);
                i++;
            }
            Doc.Close();
        }
    }
}
