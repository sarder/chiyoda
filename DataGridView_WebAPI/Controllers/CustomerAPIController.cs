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
        public int STATION_ID { get; set; }
        public string STATION_INDEX { get; set; }
        public string ENTER_SCHEDULE_TIME { get; set; }
        public string OUT_SCHEDULE_TIME { get; set; }
        public int TRANSPORT_ID { get; set; }
        public string COMPANY_NAME { get; set; }
        public string BACKGROUND_COLOR { get; set; }
        public string FONT_COLOR { get; set; }


        public results()
        {

        }
        public results(int stationId, string stationIndex, string EnterScheduleTime, string OutScheduleTime, int TransposrtId, string CompanyName, String BackgoundColor, string FontColor)
        {
            this.STATION_ID = stationId;
            this.STATION_INDEX = stationIndex;
            this.ENTER_SCHEDULE_TIME = EnterScheduleTime;
            this.OUT_SCHEDULE_TIME = OutScheduleTime;
            this.TRANSPORT_ID = TransposrtId;
            this.COMPANY_NAME = CompanyName;
            this.BACKGROUND_COLOR = BackgoundColor;
            this.FONT_COLOR = FontColor;
        }
    }
    public class CustomerAPIController : ApiController
    {


        [Route("api/CustomerAPI/GetData")]
        [HttpPost]
        public List<results> GetData()
        {
            MySqlConnection conn = WebApiConfig.conn();
            MySqlCommand query = conn.CreateCommand();

            //  query.CommandText = "select id,diagram_id,station_id,enter_schedule_time,out_schedule_time from tbl_diagram_details"; 

            //query.CommandText = "SELECT tbl_diagram_details.station_id, LOWER(SUBSTRING(mst_station.station_name, 1, 3)) as station_index,tbl_diagram_details.enter_schedule_time, tbl_diagram_details.out_schedule_time, tbl_diagram.transport_id,  mst_transport.company_name,  mst_transport.background_color as background_color, mst_transport.font_color as font_color, tbl_diagram.switching_date FROM tbl_diagram_details,tbl_diagram,mst_transport,mst_station WHERE tbl_diagram.id = tbl_diagram_details.diagram_id AND tbl_diagram.transport_id = mst_transport.id AND tbl_diagram_details.station_id = mst_station.id AND tbl_diagram.switching_date > CURDATE();";



            query.CommandText = "SELECT tbl_diagram_details.station_id, LOWER(SUBSTRING(mst_station.station_name, 1, 3)) as station_index,tbl_diagram_details.enter_schedule_time, tbl_diagram_details.out_schedule_time, tbl_diagram.transport_id,  mst_transport.company_name,  mst_transport.background_color as background_color, mst_transport.font_color as font_color, tbl_diagram.switching_date FROM tbl_diagram_details,tbl_diagram,mst_transport,mst_station WHERE tbl_diagram.id = tbl_diagram_details.diagram_id AND tbl_diagram.transport_id = mst_transport.id AND tbl_diagram_details.station_id = mst_station.id";


            // query.Parameters.AddWithValue("@staff_name", result.name);


            var results = new List<results>();
            try
            {
                conn.Open();
                MySqlDataReader fetch_query = query.ExecuteReader();
                while (fetch_query.Read())
                {
                    results.Add(new results(Convert.ToInt32(fetch_query["station_id"]), fetch_query["station_index"].ToString(), fetch_query["enter_schedule_time"].ToString(), fetch_query["out_schedule_time"].ToString(), Convert.ToInt32(fetch_query["transport_id"]), fetch_query["company_name"].ToString(), fetch_query["background_color"].ToString(), fetch_query["font_color"].ToString()));
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

            var dataTable = new DataTable();
            //try
            //{
            MySqlConnection conn = WebApiConfig.conn();
            conn.Open();

            //    MySqlCommand query = conn.CreateCommand();
            //    query.CommandText = "update tbl_excel_output set status=2 where status=1 and type =1";
            //    int isUpdaste = query.ExecuteNonQuery();

            //    if (isUpdaste > 0)
            //    {
            //        var dataSet = new DataSet();
            //        var dataAdapter = new MySqlDataAdapter { SelectCommand = InitSqlCommand("select id,diagram_id,station_id,enter_schedule_time,out_schedule_time from tbl_diagram_details where status=1 and type =1") };

            //        dataAdapter.Fill(dataSet);
            //        dataTable = dataSet.Tables[0];
            //    }

            //    conn.Close();
            //}
            //catch (Exception)
            //{

            //    throw;
            //}

            var dataSet = new DataSet();
            var dataAdapter = new MySqlDataAdapter { SelectCommand = InitSqlCommand("select id,diagram_id,station_id,enter_schedule_time,out_schedule_time from tbl_diagram_details") };

            dataAdapter.Fill(dataSet);
            dataTable = dataSet.Tables[0];

            conn.Close();

            return dataTable;

        }

        public MySqlCommand InitSqlCommand(string query)
        {
            var mySqlCommand = new MySqlCommand(query, WebApiConfig.conn());
            return mySqlCommand;
        }




        [Route("api/CustomerAPI/UpdateExcelOutputTable")]
        [HttpPost]
        public bool UpdateExcelOutputTable(string fileNameSave)
        {
            int rst = 0;
            MySqlConnection conn = WebApiConfig.conn();
            MySqlCommand query = conn.CreateCommand();

            string dateOnly = DateTime.Now.ToString("yyyy-MM-dd");

            query.CommandText = "INSERT INTO tbl_excel_output (create_datetime, begin_datetime, complete_datetime, target_date, type, status, fileserver_id, file_name) VALUES ('" + DateTime.Now.ToString() + "','" + DateTime.Now.ToString() + "','" + DateTime.Now.ToString() + "','" + dateOnly + "',1,1,192.168.0.1,'" + fileNameSave + "')";

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









        [Route("api/CustomerAPI/UploadExcel")]
        [HttpPost]
        public bool UploadExcel(string fileName)
        {


            return true;

        }












    }
}

