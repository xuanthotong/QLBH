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
    public partial class FormQMK : Form
    {
        public FormQMK()
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
            if (comboBoxTDN.SelectedItem == null && string.IsNullOrWhiteSpace(comboBoxTDN.Text))
            {
                MessageBox.Show("Vui lòng chọn tên đăng nhập!");
                return;
            }

            KetNoi = new SqlConnection(Nguon);
            string Lenh = @"SELECT TenDangNhap, MatKhau FROM TaiKhoan WHERE TenDangNhap = @TenDN";
            SqlDataAdapter da = new SqlDataAdapter(Lenh, KetNoi);
            da.SelectCommand.Parameters.AddWithValue("@TenDN", comboBoxTDN.Text);

            DataTable dt = new DataTable();
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                dataGridView1.DataSource = dt;
            }
            else
            {
                MessageBox.Show("Không tìm thấy tài khoản!");
                dataGridView1.DataSource = null;
            }
        }

        private void FormQMK_Load(object sender, EventArgs e)
        {
            KetNoi = new SqlConnection(Nguon);

            comboBoxTDN.Items.Clear();
            string Lenh = @"SELECT TenDangNhap FROM TaiKhoan";
            ThucHien = new SqlCommand(Lenh, KetNoi);

            KetNoi.Open();
            Doc = ThucHien.ExecuteReader();
            while (Doc.Read())
            {
                comboBoxTDN.Items.Add(Doc[0].ToString());
            }
            Doc.Close();
            KetNoi.Close();
        }
    }
}
