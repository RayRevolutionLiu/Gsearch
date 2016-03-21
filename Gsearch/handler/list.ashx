<%@ WebHandler Language="C#" Class="list" %>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;
using System.Data;

public class list : IHttpHandler {
    
    public void ProcessRequest (HttpContext context) {
        try
        {
            string dbtyep = context.Request.Form["dbtype"].ToString().Trim();
            string CountType = context.Request.Form["CountType"].ToString().Trim();
            string Year = context.Request.Form["Year"].ToString().Trim();
            
            SqlCommand oCmd = new SqlCommand();
            oCmd.Connection = new SqlConnection(AppConfig.DSN);
            StringBuilder sb = new StringBuilder();
            //sb.Append(@"select main_cname,main_empno,main_orgcname,main_deptcd,main_deptcdcname,ISNULL(main_viewC,0)+ISNULL(main_downloadC,0)+ISNULL(main_logC,0)+ISNULL(main_indexC,0) AS SUMt ");
            //sb.Append(@"into #tmp1 from rs_main WHERE main_empno is not null and main_empno <>''; ");
            //sb.Append(@"SELECT top 100 main_cname,main_empno,main_orgcname,main_deptcd,main_deptcdcname,SUM(SUMt) AS SUMt FROM #tmp1 GROUP BY main_cname,main_empno,main_orgcname,main_deptcd,main_deptcdcname order by SUMt DESC");
            sb.Append(@"SelectTop100Range");
            //sb.Append(@"");
            //sb.Append(@"");
            //sb.Append(@"");
            //sb.Append(@"");
            oCmd.CommandText = sb.ToString();
            oCmd.CommandType = CommandType.StoredProcedure;
            oCmd.Parameters.AddWithValue("@main_dbtype", dbtyep);
            oCmd.Parameters.AddWithValue("@CountType", CountType);
            oCmd.Parameters.AddWithValue("@Year", Year);
            SqlDataAdapter oda = new SqlDataAdapter(oCmd);
            DataTable dt = new DataTable();
            oda.Fill(dt);

            List<json> eList = new List<json>();

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    json e = new json();
                    e.cname = dt.Rows[i]["main_cname"].ToString();
                    e.empno = dt.Rows[i]["main_empno"].ToString();
                    e.orgcname = dt.Rows[i]["main_orgcname"].ToString();
                    e.main_deptcd = dt.Rows[i]["main_deptcd"].ToString();
                    e.main_deptcdcname = dt.Rows[i]["main_deptcdcname"].ToString();
                    e.SUNMt = dt.Rows[i]["SUMt"].ToString();
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
        public string cname { get; set; }
        public string empno { get; set; }
        public string orgcname { get; set; }
        public string main_deptcd { get; set; }
        public string main_deptcdcname { get; set; }
        public string SUNMt { get; set; }

    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}