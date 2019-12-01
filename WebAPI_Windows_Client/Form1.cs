using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;
using DataGridView_WebAPI.Models;
using DataGridView_WebAPI.Controllers;
using Microsoft.Win32;

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
            this.PopulateDataGridView();
            StartupProject();
        }



        public void StartupProject()
        {
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

        }

    }


}
