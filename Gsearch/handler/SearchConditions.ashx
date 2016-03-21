<%@ WebHandler Language="C#" Class="SearchConditions" %>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;
using System.Data;

public class SearchConditions : IHttpHandler {
    
    public void ProcessRequest (HttpContext context) {
        try
        {
            SqlCommand oCmd = new SqlCommand();
            oCmd.Connection = new SqlConnection(AppConfig.DSN);
            StringBuilder sb = new StringBuilder();
            sb.Append(@"SELECT main_dbtype AS value,main_dbtype AS cname ,'main_dbtype' As type_group FROM rs_main group by main_dbtype ");
            sb.Append(@"UNION ");
            sb.Append(@"select org_orgcd AS value, org_orgcd+RTRIM(LTRIM(org_abbr_chnm2)) AS cname,'orgcd' AS type_group FROM common..orgcod WHERE org_status_real = 'A'");
            //sb.Append(@"");
            //sb.Append(@"");
            //sb.Append(@"");
            //sb.Append(@"");
            //sb.Append(@"");
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
                    e.value = dt.Rows[i]["value"].ToString();
                    e.cname = dt.Rows[i]["cname"].ToString();
                    e.type_group = dt.Rows[i]["type_group"].ToString();
                    eList.Add(e);
                }
            }

            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string ans = objSerializer.Serialize(eList);  //new
            context.Response.ContentType = "application/json";
            context.Response.Write(ans);
        }
        catch(Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public class json
    {
        public string value { get; set; }
        public string cname { get; set; }
        public string type_group { get; set; }

    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}