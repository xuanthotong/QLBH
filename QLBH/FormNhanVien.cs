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
using Excel = Microsoft.Office.Interop.Excel;
namespace QLBH
{
    public partial class FormNhanVien : Form
    {
        public FormNhanVien()
        {
            InitializeComponent();
            comboBoxGioiTinh.Items.Clear();
            comboBoxGioiTinh.Items.Add("Nam");
            comboBoxGioiTinh.Items.Add("Nữ");
            comboBoxGioiTinh.SelectedIndex = 0;
        }
        string Nguon = @"Data Source=DESKTOP-FFTC51V\SQLEXPRESS;Initial Catalog=QuanLyBH;Integrated Security=True";
        string Lenh = @"";
        SqlConnection KetNoi;
        SqlCommand ThucHien;
        SqlDataReader Doc;
        SqlDataAdapter da;
        private void buttonThem_Click(object sender, EventArgs e)
        {
            //KetNoi = new SqlConnection(Nguon);
            Lenh = @"INSERT INTO NhanVien
                  (MaNV, HoTen, NgaySinh, GioiTinh, SDT, DiaChi)
                  VALUES (@MaNV,@HoTen,@NgaySinh,@GioiTinh,@SDT,@DiaChi)";
            ThucHien = new SqlCommand(Lenh,KetNoi);
            ThucHien.Parameters.Add("@MaNV", SqlDbType.Int);
            ThucHien.Parameters.Add("@HoTen", SqlDbType.NVarChar);
            ThucHien.Parameters.Add("@NgaySinh", SqlDbType.Date);
            ThucHien.Parameters.Add("@GioiTinh", SqlDbType.NVarChar);
            ThucHien.Parameters.Add("@SDT", SqlDbType.NVarChar);
            ThucHien.Parameters.Add("@DiaChi", SqlDbType.NVarChar);
            ThucHien.Parameters["@MaNV"].Value = textBoxMaNv.Text;
            ThucHien.Parameters["@HoTen"].Value = textBoxHoTen.Text;
            ThucHien.Parameters["@NgaySinh"].Value = dateTimePickerNgaySinh.Value.Date;
            ThucHien.Parameters["@GioiTinh"].Value = comboBoxGioiTinh.Text;
            ThucHien.Parameters["@SDT"].Value = textBoxSDT.Text;
            ThucHien.Parameters["@DiaChi"].Value = textBoxDiaChi.Text;
            if (comboBoxGioiTinh.SelectedIndex == -1)
            {
                MessageBox.Show("Vui lòng chọn giới tính!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            KetNoi.Open();
            ThucHien.ExecuteNonQuery();
            KetNoi.Close();
            HienThi();
        }
        void HienThi()
        {
            dataGridView.Rows.Clear();
            Lenh = @"SELECT MaNV, HoTen, NgaySinh, GioiTinh, SDT, DiaChi
                   FROM     NhanVien";
            ThucHien = new SqlCommand(Lenh, KetNoi);
            KetNoi.Open();
            Doc = ThucHien.ExecuteReader();
            int i = 0;
            while (Doc.Read())
            {
                dataGridView.Rows.Add();
                dataGridView.Rows[i].Cells[0].Value = Doc[0];
                dataGridView.Rows[i].Cells[1].Value = Doc[1];
                dataGridView.Rows[i].Cells[2].Value = Doc[2];
                dataGridView.Rows[i].Cells[3].Value = Doc[3];
                dataGridView.Rows[i].Cells[4].Value = Doc[4];
                dataGridView.Rows[i].Cells[5].Value = Doc[5];
                i++;
            }
            KetNoi.Close();
        }
        private void FormNhanVien_Load(object sender, EventArgs e)
        {
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            KetNoi = new SqlConnection(Nguon);
            HienThi();
        }

        private void buttonSua_Click(object sender, EventArgs e)
        {
            Lenh = @"UPDATE NhanVien
SET          MaNV = @MaNV, HoTen = @HoTen, NgaySinh = @NgaySinh, GioiTinh = @GioiTinh, SDT = @SDT, DiaChi = @DiaChi
WHERE  (MaNV = @Original_MaNV)";
            ThucHien = new SqlCommand(Lenh, KetNoi);
            ThucHien.Parameters.Add("@MaNV", SqlDbType.Int);
            ThucHien.Parameters.Add("@HoTen", SqlDbType.NVarChar);
            ThucHien.Parameters.Add("@NgaySinh", SqlDbType.Date);
            ThucHien.Parameters.Add("@GioiTinh", SqlDbType.NVarChar);
            ThucHien.Parameters.Add("@SDT", SqlDbType.NVarChar);
            ThucHien.Parameters.Add("@DiaChi", SqlDbType.NVarChar);
            ThucHien.Parameters["@MaNV"].Value = textBoxMaNv.Text;
            ThucHien.Parameters["@HoTen"].Value = textBoxHoTen.Text;
            ThucHien.Parameters["@NgaySinh"].Value = dateTimePickerNgaySinh.Value.Date;
            ThucHien.Parameters["@GioiTinh"].Value = comboBoxGioiTinh.Text;
            ThucHien.Parameters["@SDT"].Value = textBoxSDT.Text;
            ThucHien.Parameters["@DiaChi"].Value = textBoxDiaChi.Text;
            ThucHien.Parameters.Add("@Original_MaNV", SqlDbType.Int).Value = dataGridView.CurrentRow.Cells[0].Value;
            if (comboBoxGioiTinh.SelectedIndex == -1)
            {
                MessageBox.Show("Vui lòng chọn giới tính!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            KetNoi.Open();
            ThucHien.ExecuteNonQuery();
            KetNoi.Close();
            HienThi();
        }

        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBoxMaNv.Text = dataGridView.CurrentRow.Cells[0].Value.ToString();
            textBoxHoTen.Text = dataGridView.CurrentRow.Cells[1].Value.ToString();
            dateTimePickerNgaySinh.Value = Convert.ToDateTime(dataGridView.CurrentRow.Cells[2].Value);
            comboBoxGioiTinh.Text = dataGridView.CurrentRow.Cells[3].Value.ToString();
            textBoxSDT.Text = dataGridView.CurrentRow.Cells[4].Value.ToString();
            textBoxDiaChi.Text = dataGridView.CurrentRow.Cells[5].Value.ToString();
        }

        private void buttonXoa_Click(object sender, EventArgs e)
        {
            DialogResult D = MessageBox.Show("Bạn có muốn xóa nhân viên này không"+textBoxHoTen.Text," Chú ý",MessageBoxButtons.YesNo,MessageBoxIcon.Warning);
            if (D == DialogResult.Yes)
            {
                Lenh = @"DELETE FROM NhanVien
                     WHERE (MaNV = @Original_MaNV)";
                ThucHien = new SqlCommand(Lenh, KetNoi);
                ThucHien.Parameters.Add("@Original_MaNV", SqlDbType.Int).Value = dataGridView.CurrentRow.Cells[0].Value;
                KetNoi.Open();
                ThucHien.ExecuteNonQuery();
                KetNoi.Close();
                HienThi();
            }
            else
            {

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

        private void buttonExcel_Click(object sender, EventArgs e)
        {
            using (KetNoi = new SqlConnection(Nguon))
            {
                KetNoi.Open();
                Lenh = @"SELECT MaNV, HoTen, NgaySinh, GioiTinh, SDT, DiaChi
                   FROM     NhanVien ";
                da = new SqlDataAdapter(Lenh, KetNoi);
                DataTable dt = new DataTable();
                da.Fill(dt);
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false;
                Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Name = "DanhSachNhanVien";
                Excel.Range titleRange = worksheet.Range[
                worksheet.Cells[1, 1],
                worksheet.Cells[1, dt.Columns.Count]];
                titleRange.Merge();
                titleRange.Value = "DANH SÁCH KHÁCH HÀNG";
                titleRange.Font.Bold = true;
                titleRange.Font.Size = 16;
                titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                titleRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                titleRange.Interior.Color = Color.LightYellow;
                try
                {

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        worksheet.Cells[2, i + 1] = dt.Columns[i].ColumnName;
                    }
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 3, j + 1] = "'" + dt.Rows[i][j].ToString();
                        }
                    }
                    Excel.Range usedRange = worksheet.Range[
                       worksheet.Cells[2, 1],
                       worksheet.Cells[dt.Rows.Count + 2, dt.Columns.Count]
                   ];
                    usedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    usedRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    usedRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    usedRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    Excel.Range headerRange = worksheet.Range[
                        worksheet.Cells[2, 1],
                        worksheet.Cells[2, dt.Columns.Count]
                    ];
                    headerRange.Font.Bold = true;
                    headerRange.Interior.Color = Color.LightBlue;
                    headerRange.Font.Size = 12;
                    headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    worksheet.Columns.AutoFit();
                    worksheet.Rows.AutoFit();
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.InitialDirectory = @"D:\database\PRJ_intern\excel";
                    saveFileDialog.Title = "Chọn nơi lưu file Excel";
                    saveFileDialog.FileName = "NhanVien.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        workbook.SaveAs(saveFileDialog.FileName);
                        MessageBox.Show("Xuất Excel thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        FormNhanVien F = new FormNhanVien();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xuất Excel: " + ex.Message);
                }
                finally
                {
                    workbook.Close(false);
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                    worksheet = null;
                    workbook = null;
                    excelApp = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }
    }
}
