using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using MySql.Data.MySqlClient;

namespace Chiyoda_WebAPI
{
    public static class WebApiConfig
    {
        public static MySqlConnection conn()
        {
            string connetionString = null;
            MySqlConnection cnn;
            //connetionString = "server=localhost;database=imprest_test;username=root;password='';";
            connetionString = "server=localhost;database=imprest-share_chiyoda;username=root;password='';";
            cnn = new MySqlConnection(connetionString);
            
            return cnn;
        }
        public static void Register(HttpConfiguration config)
        {
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}
