using DataGridView_WebAPI.Models;
using System;
using System.Collections.Generic;
using System.Web.Http;
using System.Linq;
using System.Web.Http.Description;
using MySql.Data.MySqlClient;
using System.Data;

namespace DataGridView_WebAPI.Controllers
{

    public class results
    {
        public string name { get; set; }
        public string kana { get; set; }
        public int id { get; set; }
        public results()
        {

        }
        public results(string name, string kana)
        {
            this.name = name;
            this.kana = kana;
        }
    }
    public class CustomerAPIController : ApiController
    {


        [Route("api/CustomerAPI/GetCustomers")]
        [HttpPost]
        public List<results> GetCustomers(results result)
        {
            MySqlConnection conn = WebApiConfig.conn();
            MySqlCommand query = conn.CreateCommand();

            query.CommandText = "select select id,diagram_id,station_id,enter_schedule_time,out_schedule_time from tbl_diagram_details";
            query.Parameters.AddWithValue("@staff_name", result.name);

            var results = new List<results>();
            try
            {
                conn.Open();
                MySqlDataReader fetch_query = query.ExecuteReader();
                while (fetch_query.Read())
                {
                    results.Add(new results(fetch_query["staff_name"].ToString(), fetch_query["staff_kana"].ToString()));
                }
                conn.Close();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {

                //;
            }
            return results;
        }




        [Route("api/CustomerAPI/GetDataForExecl")]
        [HttpPost]
        public DataTable GetDataForExecl()
        {

           
            MySqlConnection conn = WebApiConfig.conn();
            conn.Open();

            MySqlCommand query = conn.CreateCommand();
            query.CommandText = "update tbl_excel_output set status=2 where status=1 and type =1";
            int isUpdaste = query.ExecuteNonQuery();

            var dataTable = new DataTable();

            if (isUpdaste > 0)
            {
                var dataSet = new DataSet();
                var dataAdapter = new MySqlDataAdapter { SelectCommand = InitSqlCommand("select id,diagram_id,station_id,enter_schedule_time,out_schedule_time from tbl_diagram_details where status=1 and type =1") };

                dataAdapter.Fill(dataSet);
                dataTable = dataSet.Tables[0];
            }

            conn.Close();

            return dataTable;

        }

        public MySqlCommand InitSqlCommand(string query)
        {
            var mySqlCommand = new MySqlCommand(query, WebApiConfig.conn());
            return mySqlCommand;
        }




        [Route("api/CustomerAPI/UpdateProduct")]
        [HttpPost]
        public bool UpdateProduct(results result)
        {
            int rst = 0;
            MySqlConnection conn = WebApiConfig.conn();
            MySqlCommand query = conn.CreateCommand();

            query.CommandText = "UPDATE mst_staff SET staff_name = '" + result.name + "',  staff_kana = '" + result.kana + "' WHERE id = '" + result.id + "'";

            try
            {
                conn.Open();
                rst = query.ExecuteNonQuery();
                return true;
                conn.Close();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {

                return false;
            }
            return true;

        }
    }
}

