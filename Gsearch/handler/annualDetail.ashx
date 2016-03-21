<%@ WebHandler Language="C#" Class="annualDetail" %>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;
using System.Data;

public class annualDetail : IHttpHandler {
    
    public void ProcessRequest (HttpContext context) {
        try
        {
            security secure = new security();
            string request = context.Request.Form["request"].ToString().Trim();
            string selectCountType = context.Request.Form["CountType"].ToString().Trim();
            //解開參數
            request = secure.decryptquerystring(request);
            string[] sArray = request.Split('_');
            string dbtype = sArray[0].ToString();
            string year = sArray[1].ToString();
            
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
            DataSet ds = new DataSet();
            oda.Fill(ds);

            Common common = new Common();
            DataTable dt = common.AddPercentageToDataTable(ds.Tables[0]);
                          
            List<json> eList = new List<json>();

            if (ds.Tables[0].Rows.Count > 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    json e = new json();
                    e.main_org = dt.Rows[i]["main_org"].ToString();
                    e.main_orgcname = dt.Rows[i]["main_orgcname"].ToString();
                    e.main_percentage = dt.Rows[i]["main_percentage"].ToString();
                    e.Jan = dt.Rows[i]["Jan"].ToString();
                    e.Feb = dt.Rows[i]["Feb"].ToString();
                    e.Mrz = dt.Rows[i]["Mrz"].ToString();
                    e.Apr = dt.Rows[i]["Apr"].ToString();
                    e.May = dt.Rows[i]["May"].ToString();
                    e.Jun = dt.Rows[i]["Jun"].ToString();
                    e.Jul = dt.Rows[i]["Jul"].ToString();
                    e.Aug = dt.Rows[i]["Aug"].ToString();
                    e.Sep = dt.Rows[i]["Sep"].ToString();
                    e.Oct = dt.Rows[i]["Oct"].ToString();
                    e.Nov = dt.Rows[i]["Nov"].ToString();
                    e.Dez = dt.Rows[i]["Dez"].ToString();
                    e.total = dt.Rows[i]["total"].ToString();
                    e.year = year;
                    e.dbtype = dbtype;
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
        public string main_org { get; set; }
        public string main_orgcname { get; set; }
        public string main_percentage { get; set; }     
        public string Jan { get; set; }
        public string Feb { get; set; }
        public string Mrz { get; set; }
        public string Apr { get; set; }
        public string May { get; set; }
        public string Jun { get; set; }
        public string Jul { get; set; }
        public string Aug { get; set; }
        public string Sep { get; set; }
        public string Oct { get; set; }
        public string Nov { get; set; }
        public string Dez { get; set; }
        public string total { get; set; }
        public string year { get; set; }
        public string dbtype { get; set; }

    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}