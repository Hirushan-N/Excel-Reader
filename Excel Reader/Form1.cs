using System;
using System.Linq;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Excel_Reader
{
    public partial class Reader : Form
    {
        readonly Excel.Application XLAP = new Excel.Application();
        public Reader()
        {
            InitializeComponent();
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {

            OpenFileDialog opn = new OpenFileDialog();
            opn.Title = "select file";
            opn.Filter = "Excel sheet(*.xls)|*.xls";
            opn.RestoreDirectory = true;
            opn.ShowDialog();
            txtPath.Text = opn.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show("Start");

                //SqlConnection con = new SqlConnection(connectDB.ConnectionString);
                //con.Open();


                Excel.Application xlapp = new Excel.Application();
                Excel.Workbook xlworkbook = xlapp.Workbooks.Open(txtPath.Text, "0", true, "5", "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);
                Excel.Range range = xlworksheet.UsedRange;
                int RowRange = range.Rows.Count;//12
                int ColumnRange = range.Columns.Count;//4


                int ColumnCout = 0;
                int RowCount = 0;

                String DPID = "";
                string name = "";
                string Address = "";
                String Mobile = "";
                string TP = "";
                string NIC = "";
                string Rout = "";
                decimal Basic = 0;
                DateTime RegDate = DateTime.Now;
                for (RowCount = 1; RowCount <= RowRange; RowCount++)
                {


                    for (ColumnCout = 1; ColumnCout <= ColumnRange; ColumnCout++)
                    {
                        switch (ColumnCout)
                        {
                            case 1:
                                DPID = Convert.ToString((range.Cells[RowCount, ColumnCout] as Excel.Range).Value);
                                break;
                            case 2:
                                name = Convert.ToString((range.Cells[RowCount, ColumnCout] as Excel.Range).Value);
                                break;
                            case 3:
                                Mobile = Convert.ToString((range.Cells[RowCount, ColumnCout] as Excel.Range).Value);
                                break;
                            case 4:
                                TP = Convert.ToString((range.Cells[RowCount, ColumnCout] as Excel.Range).Value);
                                break;
                            case 5:
                                Address = Convert.ToString((range.Cells[RowCount, ColumnCout] as Excel.Range).Value);
                                break;
                            case 6:
                                NIC = Convert.ToString((range.Cells[RowCount, ColumnCout] as Excel.Range).Value);
                                break;
                            case 7:
                                Rout = Convert.ToString((range.Cells[RowCount, ColumnCout] as Excel.Range).Value);
                                break;
                            case 8:
                                Basic = Convert.ToDecimal((range.Cells[RowCount, ColumnCout] as Excel.Range).Value);
                                break;
                            case 9:
                                RegDate = Convert.ToDateTime((range.Cells[RowCount, ColumnCout] as Excel.Range).Value);
                                break;

                        }
                    }






                    string sql = "INSERT INTO `delivery person` (`DP_ID`, `DP_NAME`, `MOBILE_NO`, `TP_NO`, `ADDRESS`, `NIC`, `VEHICLE_NO`, `AREA`, `BASIC_SALARY`, `REG_DATE`,`PHOTO`)" +
                                  " VALUES ('" + DPID + "', '" + name + "', '" + Mobile + "', '" + TP + "', '" + Address + "', '" + NIC + "', '" + "BDM" + (2345 + RowCount).ToString() + "', '" + Rout + "', '" + Basic + "', '" + RegDate.ToString("yyyy-MM-dd") + "',@Pic)";
                    //SqlCommand cmd = new SqlCommand(sql, con);

                    //MemoryStream ms = new MemoryStream();
                    //pictureBoxPhoto.Image.Save(ms, pictureBoxPhoto.Image.RawFormat);
                    //byte[] imgArraySav = ms.ToArray();
                    //cmd.Parameters.Add("@Pic", MySqlDbType.LongBlob).Value = imgArraySav;
                    //cmd.ExecuteNonQuery();




                }


                MessageBox.Show("Imported!");


                /*
               for (int ccnt = 1; ccnt <= rw; ccnt++)
               {
                   for (int rcnt = 1; rcnt < c1; rcnt++)
                   {
                       string str = (string)(range.Cells[rcnt][ ccnt] as Excel.Range).Value;
                       MessageBox.Show(str);
                   }
               }
                */

                xlworkbook.Close(true, null, null);
                xlapp.Quit();
                //con.Close();
            }
            catch (Exception ex)
            {

                throw ex;
            }
            MessageBox.Show("End");
        }
    }

    class connectDB
    {
        public static string ConnectionString = "datasource=localhost;port=3306;username=Hirushan;password=hirushan;database=center;";
    }
}
