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
    public partial class FormHangHoa : Form
    {
        public FormHangHoa()
        {
            InitializeComponent();
        }
        string Nguon = @"Data Source=DESKTOP-FFTC51V\SQLEXPRESS;Initial Catalog=QuanLyBH;Integrated Security=True";
        string Lenh = @"";
        SqlConnection KetNoi;
        SqlCommand ThucHien;
        SqlDataReader Doc;
        SqlDataAdapter da;
        private void buttonThem_Click(object sender, EventArgs e)
        {
            KetNoi = new SqlConnection(Nguon);
            Lenh = @"INSERT INTO HangHoa
                  (MaHH, TenHH, MaCL, DonViTinh, DonGia, SoLuong)
                 VALUES (@MaHH,@TenHH,@MaCL,@DonViTinh,@DonGia,@SoLuong)";
            ThucHien = new SqlCommand(Lenh, KetNoi);
            ThucHien.Parameters.Add("@MaHH",SqlDbType.Int);
            ThucHien.Parameters.Add("@TenHH", SqlDbType.NVarChar);
            ThucHien.Parameters.Add("@MaCL", SqlDbType.Int);
            ThucHien.Parameters.Add("@DonViTinh", SqlDbType.NVarChar);
            ThucHien.Parameters.Add("@DonGia", SqlDbType.Decimal);
            ThucHien.Parameters.Add("@SoLuong", SqlDbType.Int);
            ThucHien.Parameters["@MaHH"].Value = textBoxMaHH.Text;
            ThucHien.Parameters["@TenHH"].Value = textBoxTenHH.Text;
            ThucHien.Parameters["@MaCL"].Value = comboBoxMaCL.Text;
            ThucHien.Parameters["@DonViTinh"].Value = textBoxDVT.Text;
            ThucHien.Parameters["@DonGia"].Value = textBoxDonGia.Text;
            ThucHien.Parameters["@SoLuong"].Value = textBoxSoLuong.Text;
            KetNoi.Open();
            ThucHien.ExecuteNonQuery();
            KetNoi.Close();
            HienThi();
        }
        void HienThi()
        {
            dataGridView.Rows.Clear();
            Lenh = @"SELECT MaHH, TenHH, MaCL, DonViTinh, DonGia, SoLuong
                  FROM     HangHoa";
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

        private void FormHangHoa_Load(object sender, EventArgs e)
        {
            KetNoi = new SqlConnection(Nguon);
            HienThi();
            comboBoxMaCL.Items.Clear();
            Lenh = @"SELECT MaCL
                   FROM ChatLieu";
            ThucHien = new SqlCommand(Lenh, KetNoi);
            KetNoi.Open();
            Doc = ThucHien.ExecuteReader();
            int i = 0;
            while (Doc.Read())
            {
                comboBoxMaCL.Items.Add(Doc[0]);
                i++;
            }
            Doc.Close();
            KetNoi.Close();
        }

        private void buttonSua_Click(object sender, EventArgs e)
        {
            Lenh = @"UPDATE HangHoa
            SET          MaHH = @MaHH, TenHH = @TenHH, MaCL = @MaCL, DonViTinh = @DonViTinh, DonGia=@DonGia, SoLuong=@SoLuong
                 WHERE  (MaHH = @Original_MaHH) ";
            ThucHien = new SqlCommand(Lenh, KetNoi);
            ThucHien.Parameters.Add("@MaHH", SqlDbType.Int);
            ThucHien.Parameters.Add("@TenHH", SqlDbType.NVarChar);
            ThucHien.Parameters.Add("@MaCL", SqlDbType.Int);
            ThucHien.Parameters.Add("@DonViTinh", SqlDbType.NVarChar);
            ThucHien.Parameters.Add("@DonGia", SqlDbType.Decimal);
            ThucHien.Parameters.Add("@SoLuong", SqlDbType.Int);
            ThucHien.Parameters["@MaHH"].Value = textBoxMaHH.Text;
            ThucHien.Parameters["@TenHH"].Value = textBoxTenHH.Text;
            ThucHien.Parameters["@MaCL"].Value = comboBoxMaCL.Text;
            ThucHien.Parameters["@DonViTinh"].Value = textBoxDVT.Text;
            ThucHien.Parameters["@DonGia"].Value = textBoxDonGia.Text;
            ThucHien.Parameters["@SoLuong"].Value = textBoxSoLuong.Text;
            ThucHien.Parameters.Add("@Original_MaHH", SqlDbType.Int).Value = dataGridView.CurrentRow.Cells[0].Value;
            KetNoi.Open();
            ThucHien.ExecuteNonQuery();
            KetNoi.Close();
            HienThi();
        }

        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBoxMaHH.Text = dataGridView.CurrentRow.Cells[0].Value.ToString();
            textBoxTenHH.Text = dataGridView.CurrentRow.Cells[1].Value.ToString();
           comboBoxMaCL.Text = dataGridView.CurrentRow.Cells[2].Value.ToString();
           textBoxDVT.Text = dataGridView.CurrentRow.Cells[3].Value.ToString();
           textBoxDonGia.Text = dataGridView.CurrentRow.Cells[4].Value.ToString();
           textBoxSoLuong.Text = dataGridView.CurrentRow.Cells[5].Value.ToString();
        }

        private void buttonXoa_Click(object sender, EventArgs e)
        {
            DialogResult D = MessageBox.Show("Bạn có muốn xóa mặt hàng này không" + textBoxTenHH.Text, " Chú ý", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (D == DialogResult.Yes)
            {
                Lenh = @"DELETE FROM KhachHang
                     WHERE (MaKH = @Original_MaKH)";
                ThucHien = new SqlCommand(Lenh, KetNoi);
                ThucHien.Parameters.Add("@Original_MaHH", SqlDbType.Int).Value = dataGridView.CurrentRow.Cells[0].Value;
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
                Lenh = @"SELECT MaHH, TenHH, MaCL, DonViTinh, DonGia, SoLuong
                  FROM     HangHoa ";
                da = new SqlDataAdapter(Lenh, KetNoi);
                DataTable dt = new DataTable();
                da.Fill(dt);
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false;
                Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Name = "DanhSachHangHoa";
                Excel.Range titleRange = worksheet.Range[
                worksheet.Cells[1, 1],
                worksheet.Cells[1, dt.Columns.Count]];
                titleRange.Merge();
                titleRange.Value = "Danh sách hàng hóa";
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
                    saveFileDialog.FileName = "HangHoa.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {

                        workbook.SaveAs(saveFileDialog.FileName);
                        MessageBox.Show("Xuất Excel thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        FormHangHoa F = new FormHangHoa();
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
