<%@ WebHandler Language="C#" Class="filedownloadlist" %>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;
using System.Data;

public class filedownloadlist : IHttpHandler, System.Web.SessionState.IRequiresSessionState
{
    
    public void ProcessRequest (HttpContext context) {
        try
        {
            string selectDBIn = context.Request.Form["selectDBIn"].ToString().Trim();
            string selectCountTypeIn = context.Request.Form["selectCountTypeIn"].ToString().Trim();
            string selectCountTypeTextIn = context.Request.Form["selectCountTypeTextIn"].ToString().Trim();
            string selectYYIn = context.Request.Form["selectYYIn"].ToString().Trim();
            
            SqlCommand oCmd = new SqlCommand();
            oCmd.Connection = new SqlConnection(AppConfig.DSN);
            StringBuilder sb = new StringBuilder();
            //sb.Append(@"select main_cname,main_empno,main_orgcname,main_deptcd,main_deptcdcname,ISNULL(main_viewC,0)+ISNULL(main_downloadC,0)+ISNULL(main_logC,0)+ISNULL(main_indexC,0) AS SUMt ");
            //sb.Append(@"into #tmp1 from rs_main WHERE main_empno is not null and main_empno <>''; ");
            //sb.Append(@"SELECT top 100 main_cname,main_empno,main_orgcname,main_deptcd,main_deptcdcname,SUM(SUMt) AS SUMt FROM #tmp1 GROUP BY main_cname,main_empno,main_orgcname,main_deptcd,main_deptcdcname order by SUMt DESC");
            sb.Append(@"SelectAllRange");
            //sb.Append(@"");
            //sb.Append(@"");
            //sb.Append(@"");
            //sb.Append(@"");
            oCmd.CommandText = sb.ToString();
            oCmd.CommandType = CommandType.StoredProcedure;
            oCmd.Parameters.AddWithValue("@main_dbtype", selectDBIn);
            oCmd.Parameters.AddWithValue("@CountType", selectCountTypeIn);
            oCmd.Parameters.AddWithValue("@selectYY", selectYYIn);
            SqlDataAdapter oda = new SqlDataAdapter(oCmd);
            DataTable dt = new DataTable();
            oda.Fill(dt);

            //丟入匯出function
            ExportExcel ex = new ExportExcel();
            
            List<Part> parts = new List<Part>();
            parts.Add(new Part() { title = "資料庫", values = selectDBIn });
            parts.Add(new Part() { title = "類型", values = selectCountTypeTextIn });
            parts.Add(new Part() { title = "年分", values = selectYYIn });
            
            ex.dt = dt;
            string[] header = { "單位", "部門名稱", "部門代碼", "組別", "姓名", "工號", "職務", "職級", "年資", "信箱", "使用次數" };
            ex.header = header;
            ex.CreateExcel(parts);

            ex.CreateExcel(parts).Close();
            ex.CreateExcel(parts).Dispose();
            context.Session["downloadCSV"] = true;
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}