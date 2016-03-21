<%@ WebHandler Language="C#" Class="GenTable" %>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;
using System.Data;
using Newtonsoft.Json;

public class GenTable : IHttpHandler {

    public void ProcessRequest(HttpContext context)
    {
        try
        {
            string selected = context.Request.Form["selected"].ToString().Trim();
            string selectedOrg = context.Request.Form["selectedOrg"].ToString().Trim();
            string main_title = context.Request.Form["main_title"].ToString().Trim();
            string main_chapter = context.Request.Form["main_chapter"].ToString().Trim();
            string main_empno = context.Request.Form["main_empno"].ToString().Trim();
            string main_cname = context.Request.Form["main_cname"].ToString().Trim();
            string main_deptcd = context.Request.Form["main_deptcd"].ToString().Trim();
            string main_startdate = context.Request.Form["main_startdate"].ToString().Trim();
            string main_enddate = context.Request.Form["main_enddate"].ToString().Trim();
            string sortcolumn = context.Request.Form["sortcolumn"].ToString().Trim();
            string orderBy = context.Request.Form["orderBy"].ToString().Trim();

            string pageSize = context.Request.Form["pageSize"].ToString().Trim();
            //string frameHeight = context.Request.Form["frameHeight"].ToString().Trim();
            string flag = context.Request.Form["flag"].ToString().Trim();
            string Export = context.Request.Form["Export"].ToString().Trim();

            int ENDs = int.Parse(pageSize) * int.Parse(flag);
            int STARTs = ENDs - int.Parse(pageSize) + 1;

            SqlCommand oCmd = new SqlCommand();
            oCmd.Connection = new SqlConnection(AppConfig.DSN);
            StringBuilder sb = new StringBuilder();
            sb.Append(@"GenTable");
            //sb.Append(@"");
            oCmd.CommandText = sb.ToString();
            oCmd.CommandType = CommandType.StoredProcedure;
            oCmd.Parameters.AddWithValue("@selected", selected);
            oCmd.Parameters.AddWithValue("@selectedOrg", selectedOrg);
            oCmd.Parameters.AddWithValue("@main_title", main_title);
            oCmd.Parameters.AddWithValue("@main_chapter", main_chapter);
            oCmd.Parameters.AddWithValue("@main_empno", main_empno);
            oCmd.Parameters.AddWithValue("@main_cname", main_cname);
            oCmd.Parameters.AddWithValue("@main_deptcd", main_deptcd);
            oCmd.Parameters.AddWithValue("@main_startdate", main_startdate);
            oCmd.Parameters.AddWithValue("@main_enddate", main_enddate);

            oCmd.Parameters.AddWithValue("@START", STARTs);
            oCmd.Parameters.AddWithValue("@END", ENDs);

            oCmd.Parameters.AddWithValue("@sortcolumn", sortcolumn);
            oCmd.Parameters.AddWithValue("@orderBy", orderBy);
            oCmd.Parameters.AddWithValue("@Export", Export);
            SqlDataAdapter oda = new SqlDataAdapter(oCmd);
            DataSet ds = new DataSet();
            oda.Fill(ds);

            List<json> eList = new List<json>();

            if (ds.Tables[0].Rows.Count > 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    json e = new json();
                    e.main_ip = ds.Tables[0].Rows[i]["main_ip"].ToString();
                    e.main_dbtype = ds.Tables[0].Rows[i]["main_dbtype"].ToString();
                    e.main_viewC = ds.Tables[0].Rows[i]["main_viewC"].ToString();
                    e.main_downloadC = ds.Tables[0].Rows[i]["main_downloadC"].ToString();
                    e.main_logC = ds.Tables[0].Rows[i]["main_logC"].ToString();
                    e.main_indexC = ds.Tables[0].Rows[i]["main_indexC"].ToString();
                    e.main_datetime = ds.Tables[0].Rows[i]["main_datetime"].ToString().Trim() == "" ? "" : Convert.ToDateTime(ds.Tables[0].Rows[i]["main_datetime"].ToString()).ToString("yyyy/MM/dd");
                    e.main_title = ds.Tables[0].Rows[i]["main_title"].ToString();
                    e.main_chapter = ds.Tables[0].Rows[i]["main_chapter"].ToString();
                    e.main_empno = ds.Tables[0].Rows[i]["main_empno"].ToString();
                    e.main_cname = ds.Tables[0].Rows[i]["main_cname"].ToString();
                    e.main_org = ds.Tables[0].Rows[i]["main_org"].ToString();
                    e.main_email = ds.Tables[0].Rows[i]["main_email"].ToString();
                    e.main_orgcname = ds.Tables[0].Rows[i]["main_orgcname"].ToString();
                    e.main_deptcd = ds.Tables[0].Rows[i]["main_deptcd"].ToString();
                    e.main_deptcdcname = ds.Tables[0].Rows[i]["main_deptcdcname"].ToString();
                    e.totalcount = ds.Tables[1].Rows[0]["quity"].ToString();
                    //e.value = ds.Tables[0].Rows[i]["value"].ToString();
                    eList.Add(e);
                }

                string ans = JsonConvert.SerializeObject(eList);
                //System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                //string ans = objSerializer.Serialize(eList);  //new
                context.Response.ContentType = "application/json";
                context.Response.Write(ans);
            }
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }


    public class json
    {
        public string main_ip { get; set; }
        public string main_dbtype { get; set; }
        public string main_viewC { get; set; }
        public string main_downloadC { get; set; }
        public string main_logC { get; set; }
        public string main_indexC { get; set; }
        public string main_datetime { get; set; }
        public string main_title { get; set; }
        public string main_chapter { get; set; }
        public string main_empno { get; set; }
        public string main_cname { get; set; }
        public string main_email { get; set; }
        public string main_org { get; set; }
        public string main_orgcname { get; set; }
        public string main_deptcd { get; set; }
        public string main_deptcdcname { get; set; }
        public string totalcount { get; set; }
    }
    
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}