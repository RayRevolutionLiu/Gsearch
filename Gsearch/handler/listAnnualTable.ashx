<%@ WebHandler Language="C#" Class="listAnnualTable" %>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;
using System.Data;

public class listAnnualTable : IHttpHandler {
    
    public void ProcessRequest (HttpContext context) {
        try
        {
            string dbtype = context.Request.Form["dbtype"].ToString().Trim();
            SqlCommand oCmd = new SqlCommand();
            oCmd.Connection = new SqlConnection(AppConfig.DSN);
            StringBuilder sb = new StringBuilder();
            sb.Append(@"SELECT ");
            sb.Append(@"statistics_dbtype,");
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
            sb.Append(@"(Select statistics_dbtype,statistics_num,MONTH(statistics_datetime) as TMonth from ");
            sb.Append(@"rs_statistics WHERE YEAR(statistics_datetime) = @dbtype) source ");
            sb.Append(@"PIVOT (");
            sb.Append(@" SUM(statistics_num) ");
            sb.Append(@" FOR TMonth");
            sb.Append(@" IN ( [1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12] )");
            sb.Append(@" ) AS pvtMonth;");
            sb.Append(@"SELECT ");
            sb.Append(@"statistics_dbtype,");
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
            sb.Append(@"[12] AS Dez into #tmp2 FROM ");
            sb.Append(@"(Select statistics_dbtype,statistics_parentid,MONTH(statistics_datetime) as TMonth from ");
            sb.Append(@"rs_statistics WHERE YEAR(statistics_datetime) = @dbtype) source ");
            sb.Append(@"PIVOT (");
            sb.Append(@" MAX(statistics_parentid) ");
            sb.Append(@" FOR TMonth");
            sb.Append(@" IN ( [1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12] )");
            sb.Append(@" ) AS pvtMonth;");
            sb.Append(@"select *,(ISNULL(Jan,0)+ISNULL(Feb,0)+");
            sb.Append(@"ISNULL(Mrz,0)+ISNULL(Apr,0)+ISNULL(May,0)+ISNULL(Jun,0)+");
            sb.Append(@"ISNULL(Jul,0)+ISNULL(Aug,0)+ISNULL(Sep,0)+ISNULL(Oct,0)+");
            sb.Append(@"ISNULL(Nov,0)+ISNULL(Dez,0)) AS total FROM #tmp1;");
            sb.Append(@"SELECT *,statistics_dbtype+'_'+@dbtype AS total FROM #tmp2");
            sb.Append(@"");
            oCmd.CommandText = sb.ToString();
            oCmd.CommandType = CommandType.Text;
            oCmd.Parameters.AddWithValue("@dbtype", dbtype);
            SqlDataAdapter oda = new SqlDataAdapter(oCmd);
            DataSet ds = new DataSet();
            oda.Fill(ds);

            List<Cover> cList = new List<Cover>();
            security secure = new security();
            
            if (ds.Tables.Count > 0)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    Cover Ce = new Cover();
                    
                    for (int j = 0; j < ds.Tables.Count; j++)
                    {
                        List<json> eList = new List<json>();
                        
                        //因為Tables[0]跟Tables[1]語法完全一樣,所以只要判斷一個table即可
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            json e = new json();
                            e.statistics_dbtype = ds.Tables[0].Rows[i]["statistics_dbtype"].ToString();
                            e.Jan = ds.Tables[j].Rows[i]["Jan"].ToString();
                            e.Feb = ds.Tables[j].Rows[i]["Feb"].ToString();
                            e.Mrz = ds.Tables[j].Rows[i]["Mrz"].ToString();
                            e.Apr = ds.Tables[j].Rows[i]["Apr"].ToString();
                            e.May = ds.Tables[j].Rows[i]["May"].ToString();
                            e.Jun = ds.Tables[j].Rows[i]["Jun"].ToString();
                            e.Jul = ds.Tables[j].Rows[i]["Jul"].ToString();
                            e.Aug = ds.Tables[j].Rows[i]["Aug"].ToString();
                            e.Sep = ds.Tables[j].Rows[i]["Sep"].ToString();
                            e.Oct = ds.Tables[j].Rows[i]["Oct"].ToString();
                            e.Nov = ds.Tables[j].Rows[i]["Nov"].ToString();
                            e.Dez = ds.Tables[j].Rows[i]["Dez"].ToString();
                            if (j > 0)
                            {
                                e.total = secure.encryptquerystring(ds.Tables[j].Rows[i]["total"].ToString());
                            }
                            else
                            {
                                e.total = ds.Tables[j].Rows[i]["total"].ToString();
                            }
                            eList.Add(e);
                        }
                        if (j == 0)
                        {
                            Ce.NumberTable = eList;
                        }
                        else
                        {
                            Ce.UidTable = eList;
                        }
                    }
                    cList.Add(Ce);
                }
            }

            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string ans = objSerializer.Serialize(cList);  //new
            context.Response.ContentType = "application/json";
            context.Response.Write(ans);
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public class Cover
    {
        public List<json> NumberTable { get; set; }
        public List<json> UidTable { get; set; }
    }

    public class json
    {
        public string statistics_dbtype { get; set; }
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

    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}