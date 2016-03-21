<%@ WebHandler Language="C#" Class="getYYgroupBy" %>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;
using System.Data;

public class getYYgroupBy : IHttpHandler {

    public void ProcessRequest(HttpContext context)
    {
        try
        {
            SqlCommand oCmd = new SqlCommand();
            oCmd.Connection = new SqlConnection(AppConfig.DSN);
            StringBuilder sb = new StringBuilder();
            sb.Append(@"select YEAR(statistics_datetime) AS yyyy FROM  rs_statistics GROUP BY YEAR(statistics_datetime) ORDER BY YEAR(statistics_datetime) DESC ");
            oCmd.CommandText = sb.ToString();
            oCmd.CommandType = CommandType.Text;
            SqlDataAdapter oda = new SqlDataAdapter(oCmd);
            DataTable dt = new DataTable();
            oda.Fill(dt);

            List<json> eList = new List<json>();

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    json e = new json();
                    e.value = dt.Rows[i]["yyyy"].ToString();
                    e.cname = dt.Rows[i]["yyyy"].ToString();
                    //e.type_group = dt.Rows[i]["type_group"].ToString();
                    eList.Add(e);
                }
            }

            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string ans = objSerializer.Serialize(eList);  //new
            context.Response.ContentType = "application/json";
            context.Response.Write(ans);
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public class json
    {
        public string value { get; set; }
        public string cname { get; set; }

    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}