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

            SqlConnection cnn;
            string connectionString = null;
            string sql = null;
            string data = null;
            int i = 0;
            int j = 0;

            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            System.Data.DataTable ds = getData;

            for (i = 0; i <= ds.Rows.Count - 1; i++)
            {
                for (j = 0; j <= ds.Columns.Count - 1; j++)
                {
                    data = ds.Rows[i].ItemArray[j].ToString();
                    xlWorkSheet.Cells[i + 1, j + 1] = data;
                    xlWorkSheet.Cells[i + 1, j + 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
                }
            }




            string root = @"D:\chiyoda\";
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

            MessageBox.Show("Excel file created , you can find the file "+ root + xcelFileName + ".xlsx");

        }

        //Execl work related method
        private void btnExcel_Click(object sender, EventArgs e)
        {

            getDataFromAPI();

            /*
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Range chartRange;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //add data 
            xlWorkSheet.Cells[4, 2] = "";
            xlWorkSheet.Cells[4, 3] = "Student1";
            xlWorkSheet.Cells[4, 4] = "Student2";
            xlWorkSheet.Cells[4, 5] = "Student3";

            xlWorkSheet.Cells[5, 2] = "Term1";
            xlWorkSheet.Cells[5, 3] = "80";
            xlWorkSheet.Cells[5, 4] = "65";
            xlWorkSheet.Cells[5, 5] = "45";

            xlWorkSheet.Cells[6, 2] = "Term2";
            xlWorkSheet.Cells[6, 3] = "78";
            xlWorkSheet.Cells[6, 4] = "72";
            xlWorkSheet.Cells[6, 5] = "60";

            xlWorkSheet.Cells[7, 2] = "Term3";
            xlWorkSheet.Cells[7, 3] = "82";
            xlWorkSheet.Cells[7, 4] = "80";
            xlWorkSheet.Cells[7, 5] = "65";

            xlWorkSheet.Cells[8, 2] = "Term4";
            xlWorkSheet.Cells[8, 3] = "75";
            xlWorkSheet.Cells[8, 4] = "82";
            xlWorkSheet.Cells[8, 5] = "68";

            xlWorkSheet.Cells[9, 2] = "Total";
            xlWorkSheet.Cells[9, 3] = "315";
            xlWorkSheet.Cells[9, 4] = "299";
            xlWorkSheet.Cells[9, 5] = "238";

            xlWorkSheet.get_Range("b2", "e3").Merge(false);

            chartRange = xlWorkSheet.get_Range("b2", "e3");
            chartRange.FormulaR1C1 = "MARK LIST";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
            chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            chartRange.Font.Size = 20;

            chartRange = xlWorkSheet.get_Range("b4", "e4");
            chartRange.Font.Bold = true;
            chartRange = xlWorkSheet.get_Range("b9", "e9");
            chartRange.Font.Bold = true;

            chartRange = xlWorkSheet.get_Range("b2", "e9");
            chartRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic);

            xlWorkBook.SaveAs("d:\\csharp.net-informations.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlApp);
            releaseObject(xlWorkBook);
            releaseObject(xlWorkSheet);

            MessageBox.Show("File created !");*/

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
