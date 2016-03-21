<%@ WebHandler Language="C#" Class="filedownloadAnnual" %>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;
using System.Data;

public class filedownloadAnnual : IHttpHandler, System.Web.SessionState.IRequiresSessionState
{
    
    public void ProcessRequest (HttpContext context) {
        try
        {
            string uid = context.Request.Form["uid"].ToString().Trim();

            security secure = new security();
            uid = secure.decryptquerystring(uid);
            string[] sArray = uid.Split('_');
            string dbtype = sArray[0].ToString();
            string year = sArray[1].ToString();
            
            string selectCountType = context.Request.Form["selectCountType"].ToString().Trim();
            string selectCountTypeText = context.Request.Form["selectCountTypeText"].ToString().Trim();

            SqlCommand oCmd = new SqlCommand();
            oCmd.Connection = new SqlConnection(AppConfig.DSN);
            StringBuilder sb = new StringBuilder();
            if (selectCountType.Trim() == "")
            {
                //抓出全部
                sb.Append(@"SELECT ISNULL(main_org,'未知')AS main_org,ISNULL(main_orgcname,'未知資料') AS main_orgcname,");
                sb.Append(@"MONTH(main_datetime) AS monthC,");
                sb.Append(@"ISNULL(SUM(main_viewC),0) AS viewC,");
                sb.Append(@"ISNULL(SUM(main_downloadC),0) AS downloadC,");
                sb.Append(@"ISNULL(SUM(main_logC),0) AS logC,");
                sb.Append(@"ISNULL(SUM(main_indexC),0) AS indexC ");
                sb.Append(@"into #tmpInside FROM rs_main ");
                sb.Append(@"WHERE YEAR(main_datetime) = @year AND main_dbtype = @dbtype ");
                sb.Append(@"group by main_org,main_orgcname,MONTH(main_datetime);");
            }
            else
            {
                sb.Append(@"SELECT ISNULL(main_org,'未知')AS main_org,ISNULL(main_orgcname,'未知資料') AS main_orgcname,");
                sb.Append(@"MONTH(main_datetime) AS monthC,");
                sb.Append(@"ISNULL(SUM(" + selectCountType + "),0) AS " + selectCountType + " ");
                sb.Append(@"into #tmpInside FROM rs_main ");
                sb.Append(@"WHERE YEAR(main_datetime) = @year AND main_dbtype = @dbtype ");
                sb.Append(@"group by main_org,main_orgcname,MONTH(main_datetime);");
            }
            sb.Append(@"SELECT ");
            sb.Append(@"ISNULL(main_org,'未知')AS main_org,ISNULL(main_orgcname,'未知資料') AS main_orgcname,");
            sb.Append(@"[1] AS Jan,");
            sb.Append(@"[2] AS Feb,");
            sb.Append(@"[3] AS Mrz,");
            sb.Append(@"[4] AS Apr,");
            sb.Append(@"[5] AS May,");
            sb.Append(@"[6] AS Jun,");
            sb.Append(@"[7] AS Jul,");
            sb.Append(@"[8] AS Aug,");
            sb.Append(@"[9] AS Sep,");
            sb.Append(@"[10] AS Oct,");
            sb.Append(@"[11] AS Nov,");
            sb.Append(@"[12] AS Dez into #tmp1 FROM ");
            sb.Append(@" ( ");
            if (selectCountType.Trim() == "")
            {
                sb.Append(@"SELECT main_org,main_orgcname,monthC,viewC+downloadC+logC+indexC AS Allcol FROM #tmpInside ");
            }
            else
            {
                sb.Append(@"SELECT main_org,main_orgcname,monthC," + selectCountType + " AS Allcol FROM #tmpInside ");
            }
            sb.Append(@" ) ");
            sb.Append(@" source ");
            sb.Append(@"PIVOT (");
            sb.Append(@" SUM(Allcol) ");
            sb.Append(@" FOR monthC");
            sb.Append(@" IN ( [1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12] )");
            sb.Append(@" ) AS pvtMonth;");

            sb.Append(@"select main_org,main_orgcname,ISNULL(Jan,0) AS Jan,");
            sb.Append(@"ISNULL(Feb,0) AS Feb,");
            sb.Append(@"ISNULL(Mrz,0) AS Mrz,");
            sb.Append(@"ISNULL(Apr,0) AS Apr,");
            sb.Append(@"ISNULL(May,0) AS May,");
            sb.Append(@"ISNULL(Jun,0) AS Jun,");
            sb.Append(@"ISNULL(Jul,0) AS Jul,");
            sb.Append(@"ISNULL(Aug,0) AS Aug,");
            sb.Append(@"ISNULL(Sep,0) AS Sep,");
            sb.Append(@"ISNULL(Oct,0) AS Oct,");
            sb.Append(@"ISNULL(Nov,0) AS Nov,");
            sb.Append(@"ISNULL(Dez,0) AS Dez,");
            sb.Append(@"(ISNULL(Jan,0)+ISNULL(Feb,0)+");
            sb.Append(@"ISNULL(Mrz,0)+ISNULL(Apr,0)+ISNULL(May,0)+ISNULL(Jun,0)+");
            sb.Append(@"ISNULL(Jul,0)+ISNULL(Aug,0)+ISNULL(Sep,0)+ISNULL(Oct,0)+");
            sb.Append(@"ISNULL(Nov,0)+ISNULL(Dez,0)) AS total FROM #tmp1 ORDER BY total DESC;");
            
            oCmd.CommandText = sb.ToString();
            oCmd.CommandType = CommandType.Text;
            oCmd.Parameters.AddWithValue("@dbtype", dbtype);
            oCmd.Parameters.AddWithValue("@year", year);
            SqlDataAdapter oda = new SqlDataAdapter(oCmd);
            DataTable dt = new DataTable();
            oda.Fill(dt);

            Common common = new Common();
            dt = common.AddPercentageToDataTable(dt);
            
            //丟入匯出function
            ExportExcel ex = new ExportExcel();

            List<Part> parts = new List<Part>();
            parts.Add(new Part() { title = "資料庫", values = dbtype });
            parts.Add(new Part() { title = "類型", values = selectCountTypeText });
            parts.Add(new Part() { title = "年分", values = year });
            ex.dt = dt;
            string[] header = { "單位代碼", "單位名稱", "百分比", "1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "小計" };
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