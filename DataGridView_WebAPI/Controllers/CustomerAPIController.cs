using DataGridView_WebAPI.Models;
using System;
using System.Collections.Generic;
using System.Web.Http;
using System.Linq;
using System.Web.Http.Description;
using MySql.Data.MySqlClient;


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

            query.CommandText = "select id,staff_name,staff_kana from mst_staff where staff_name=@staff_name";
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

