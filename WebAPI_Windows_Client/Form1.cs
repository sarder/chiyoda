using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;
using DataGridView_WebAPI.Models;
using DataGridView_WebAPI.Controllers;
using Microsoft.Win32;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Application = System.Windows.Forms.Application;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Data;
using System.Data.SqlClient;
using Newtonsoft.Json;
using System.IO;

namespace WebAPI_Windows_Client
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string apiUrl = "http://localhost:26404/api/CustomerAPI";

        private void Form1_Load(object sender, EventArgs e)
        {
            // this.PopulateDataGridView();
            //  StartupProject();

            timerCheckTime.Enabled = true;
        }



        public void StartupProject()
        {

            //software copy inside of c drive: task
            RegistryKey registryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            registryKey.SetValue("Chiyoda", Application.ExecutablePath.ToString());
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            this.PopulateDataGridView();
        }

        private void PopulateDataGridView()
        {

            object input = new
            {
                Name = txtName.Text.Trim(),
            };
            string inputJson = (new JavaScriptSerializer()).Serialize(input);
            WebClient client = new WebClient();
            client.Headers["Content-type"] = "application/json";
            client.Encoding = Encoding.UTF8;
            string json = client.UploadString(apiUrl + "/GetCustomers", inputJson);

            dataGridView1.DataSource = (new JavaScriptSerializer()).Deserialize<List<CustomerModel>>(json);
        }




        private void button1_Click(object sender, EventArgs e)
        {

            updateCustomer();
            //update code
        }


        private async void updateCustomer()
        {


            results p = new results();
            p.kana = "Rolex";
            p.name = "Watch";
            p.id = 20;


            string inputJson = (new JavaScriptSerializer()).Serialize(p);
            WebClient client = new WebClient();
            client.Headers["Content-type"] = "application/json";
            client.Encoding = Encoding.UTF8;
            string json = client.UploadString(apiUrl + "/UpdateProduct", inputJson);

            MessageBox.Show("checked done");

        }

        private void timerCheckTime_Tick(object sender, EventArgs e)
        {
            updateCustomer();
        }




        #region MyRegion


        public void getDataFromAPI()
        {

            //new code ************************************************************************************************************

            WebClient client = new WebClient();
            client.Headers["Content-type"] = "application/json";
            client.Encoding = Encoding.UTF8;
            string json = client.UploadString(apiUrl + "/GetDataForExecl", "");

            System.Data.DataTable getData = (System.Data.DataTable)JsonConvert.DeserializeObject(json, (typeof(System.Data.DataTable)));

            string data;
            int i = 0;
            int j = 0;

            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Shapes.AddPicture(@"D:\Japan\chiyoda\WebAPI_Windows_Client\img\ch.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 1550, 5, 270, 180);


            System.Data.DataTable ds = getData;



            //static header section start

            xlWorkSheet.get_Range("a1", "v2").Merge(false);
            Microsoft.Office.Interop.Excel.Range chartRange;
            chartRange = xlWorkSheet.get_Range("a1", "v2");
            chartRange.FormulaR1C1 = "千代田工業㈱";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 20;


            xlWorkSheet.get_Range("a3", "v4").Merge(false);
            chartRange = xlWorkSheet.get_Range("a3", "v4");
            chartRange.FormulaR1C1 = "豊明工場ST表";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 20;


            xlWorkSheet.get_Range("a10", "v11").Merge(false);
            chartRange = xlWorkSheet.get_Range("a10", "v11");
            chartRange.FormulaR1C1 = "2019年 11月29日～";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 20;

            xlWorkSheet.get_Range("y1", "af2").Merge(false);
            chartRange = xlWorkSheet.get_Range("y1", "af2");
            chartRange.FormulaR1C1 = "ST配置図";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 10;



            xlWorkSheet.get_Range("bq1", "dm1").Merge(false);
            chartRange = xlWorkSheet.get_Range("bq1", "dm1");
            chartRange.FormulaR1C1 = "月度変化点";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 10;
            xlWorkSheet.Range["bq1", "dm1"].Borders.Color = Color.Black;

          

            xlWorkSheet.get_Range("bq2", "dm3").Merge(false);
            chartRange = xlWorkSheet.get_Range("bq2", "dm3");
           // chartRange.FormulaR1C1 = "月度変化点";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 15;
            xlWorkSheet.Range["bq2", "dm3"].Borders.Color = Color.Black;



            xlWorkSheet.get_Range("bq4", "dm5").Merge(false);
            chartRange = xlWorkSheet.get_Range("bq4", "dm5");
            // chartRange.FormulaR1C1 = "月度変化点";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 15;
            xlWorkSheet.Range["bq4", "dm5"].Borders.Color = Color.Black;



            xlWorkSheet.get_Range("bq6", "dm7").Merge(false);
            chartRange = xlWorkSheet.get_Range("bq6", "dm7");
            // chartRange.FormulaR1C1 = "月度変化点";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 15;
            xlWorkSheet.Range["bq6", "dm7"].Borders.Color = Color.Black;



            xlWorkSheet.get_Range("bq8", "dm9").Merge(false);
            chartRange = xlWorkSheet.get_Range("bq8", "dm9");
            // chartRange.FormulaR1C1 = "月度変化点";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 15;
            xlWorkSheet.Range["bq8", "dm9"].Borders.Color = Color.Black;




            xlWorkSheet.get_Range("bq10", "dm11").Merge(false);
            chartRange = xlWorkSheet.get_Range("bq10", "dm11");
            // chartRange.FormulaR1C1 = "月度変化点";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 15;
            xlWorkSheet.Range["bq10", "dm11"].Borders.Color = Color.Black;



            xlWorkSheet.get_Range("eh1", "ew1").Merge(false);
            chartRange = xlWorkSheet.get_Range("eh1", "ew1");
            chartRange.FormulaR1C1 = "生産管理部工務改善課";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 10;


            xlWorkSheet.get_Range("eh2", "ew2").Merge(false);
            chartRange = xlWorkSheet.get_Range("eh2", "ew2");
            chartRange.FormulaR1C1 = "物流改善係";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Font.Size = 10;





            //static header section end


            xlWorkSheet.get_Range("a15", "b50").Merge(false);
            chartRange = xlWorkSheet.get_Range("a15", "b50");
            chartRange.FormulaR1C1 = "昼勤";
            chartRange.HorizontalAlignment = 2;
            chartRange.VerticalAlignment = 2;
            chartRange.Font.Size = 10;





            xlApp.ActiveWindow.DisplayGridlines = false;


            //data get from database using json

          /*  for (i = 0; i <= ds.Rows.Count - 1; i++)
            {
                for (j = 0; j <= ds.Columns.Count - 1; j++)
                {
                    data = ds.Rows[i].ItemArray[j].ToString();
                    xlWorkSheet.Cells[i + 60, j + 1] = data;
                    //xlWorkSheet.Cells[i + 1, j + 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
                }
            }
            */



            //dynamic section start here 

            Microsoft.Office.Interop.Excel.Range formatRange;
            formatRange = xlWorkSheet.get_Range("a15", "ew50");
            formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous,
            Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic,
            Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic);


         





            //dynamic section end here 

            xlWorkSheet.Cells.ColumnWidth = 1;
         
            string root = @"C:\chiyoda\";
            if (!Directory.Exists(root))
            {
                Directory.CreateDirectory(root);
            }


            string datetime = DateTime.Now.ToString();
            string xcelFileName = ReplaceHelper.DateTimeStringBuilder(datetime);


            xlWorkBook.SaveAs(root + xcelFileName + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file " + root + xcelFileName + ".xlsx");

        }

        //Execl work related method
        private void btnExcel_Click(object sender, EventArgs e)
        {
            getDataFromAPI();
        }


        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        #endregion


    }


}
