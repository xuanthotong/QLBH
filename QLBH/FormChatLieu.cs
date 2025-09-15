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
    public partial class FormChatLieu : Form
    {
        public FormChatLieu()
        {
            InitializeComponent();
        }
        string Nguon = @"Data Source=DESKTOP-FFTC51V\SQLEXPRESS;Initial Catalog=QuanLyBH;Integrated Security=True";
        string Lenh = @"";
        SqlConnection KetNoi;
        SqlCommand ThucHien;
        SqlDataReader Doc;
        SqlDataAdapter da;
        void HienThi()
        {
            dataGridView1.Rows.Clear();
            Lenh = @"SELECT MaCL, TenCL
                     FROM     ChatLieu";
            ThucHien = new SqlCommand(Lenh, KetNoi);
            KetNoi.Open();
            Doc = ThucHien.ExecuteReader();
            int i = 0;
            while (Doc.Read())
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Cells[0].Value = Doc[0];
                dataGridView1.Rows[i].Cells[1].Value = Doc[1];
                i++;
            }
            KetNoi.Close();
        }
        private void FormChatLieu_Load_1(object sender, EventArgs e)
        {
            KetNoi = new SqlConnection(Nguon);
            HienThi();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBoxMaCL1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBoxTenCL1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
        }

        private void buttonThem1_Click_1(object sender, EventArgs e)
        {
            KetNoi = new SqlConnection(Nguon);
            Lenh = @"INSERT INTO ChatLieu
                  (MaCL, TenCL)
                  VALUES (@MaCL,@TenCL)";
            ThucHien = new SqlCommand(Lenh, KetNoi);
            ThucHien.Parameters.Add("@MaCL", SqlDbType.Int);
            ThucHien.Parameters.Add("@TenCL", SqlDbType.NVarChar);
            ThucHien.Parameters["@MaCL"].Value = textBoxMaCL1.Text;
            ThucHien.Parameters["@TenCL"].Value = textBoxTenCL1.Text;
            KetNoi.Open();
            ThucHien.ExecuteNonQuery();
            KetNoi.Close();
            HienThi();
        }

        private void buttonSua1_Click_1(object sender, EventArgs e)
        {
            Lenh = @"UPDATE ChatLieu
                 SET          MaCL = @MaCL, TenCL = @TenCL
               WHERE  (MaCL = @Original_MaCL) ";
            ThucHien = new SqlCommand(Lenh, KetNoi);
            ThucHien.Parameters.Add("@MaCL", SqlDbType.Int);
            ThucHien.Parameters.Add("@TenCL", SqlDbType.NVarChar);
            ThucHien.Parameters["@MaCL"].Value = textBoxMaCL1.Text;
            ThucHien.Parameters["@TenCL"].Value = textBoxTenCL1.Text;
            ThucHien.Parameters.Add("@Original_MaCL", SqlDbType.Int).Value = dataGridView1.CurrentRow.Cells[0].Value;
            KetNoi.Open();
            ThucHien.ExecuteNonQuery();
            KetNoi.Close();
            HienThi();
        }

        private void buttonXoa1_Click(object sender, EventArgs e)
        {
            DialogResult D = MessageBox.Show("Bạn có muốn xóa nhân viên này không" + textBoxMaCL1.Text, " Chú ý", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (D == DialogResult.Yes)
            {
                Lenh = @"DELETE FROM ChatLieu
                     WHERE (MaCL = @Original_MaCL)";
                ThucHien = new SqlCommand(Lenh, KetNoi);
                ThucHien.Parameters.Add("@Original_MaCL", SqlDbType.Int).Value = dataGridView1.CurrentRow.Cells[0].Value;
                KetNoi.Open();
                ThucHien.ExecuteNonQuery();
                KetNoi.Close();
                HienThi();
            }
            else
            {

            }
        }

        private void buttonThoat1_Click(object sender, EventArgs e)
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

        private void buttonExcel1_Click(object sender, EventArgs e)
        {
            using (KetNoi = new SqlConnection(Nguon))
            {
                KetNoi.Open();
                Lenh = @"SELECT MaCL, TenCL
                    FROM     ChatLieu ";
                da = new SqlDataAdapter(Lenh, KetNoi);
                DataTable dt = new DataTable();
                da.Fill(dt);
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false;
                Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Name = "ChatLieu";
                Excel.Range titleRange = worksheet.Range[
                worksheet.Cells[1, 1],
                worksheet.Cells[1, dt.Columns.Count]
                ];
                titleRange.Merge();
                titleRange.Value = "Chất Liệu";
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
                            worksheet.Cells[i + 3, j + 1] = dt.Rows[i][j].ToString();
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
                    saveFileDialog.Title = "Chọn nơi lưu file Excel";
                    saveFileDialog.FileName = "ChatLieu.xlsx";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        workbook.SaveAs(saveFileDialog.FileName);
                        MessageBox.Show("Xuất Excel thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        FormChatLieu F = new FormChatLieu();
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
