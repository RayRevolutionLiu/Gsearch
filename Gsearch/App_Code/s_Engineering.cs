using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using System.Globalization;
using System.Net;

/// <summary>
/// s_Engineering 的摘要描述
/// </summary>
public class s_Engineering
{
	public s_Engineering()
	{
		//
		// TODO: 在這裡新增建構函式邏輯
		//
	}

    public int Execute(DataTable dt, string strGUID)
    {
        int rowcount = dt.Rows.Count;
        //int rowcount = 13;
        bool c = true;
        int realImportNum = 0;

        StringBuilder sb = new StringBuilder();
        sb.Append(@"INSERT INTO rs_main (");
        sb.Append(@"main_parentid,main_ip,main_dbtype,main_title,");
        sb.Append(@"main_downloadC,main_datetime,");
        sb.Append(@"main_createdate");
        //sb.Append(@"");
        //sb.Append(@"");
        sb.Append(@" ) ");
        sb.Append(@" VALUES (");
        sb.Append(@"@main_parentid,@main_ip,@main_dbtype,@main_title,");
        sb.Append(@"@main_downloadC,@main_datetime,");
        sb.Append(@"@main_createdate");
        //sb.Append(@"");
        sb.Append(@")");

        SqlConnection oConn = new SqlConnection(AppConfig.DSN);
        oConn.Open();
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = oConn;
        SqlTransaction myTrans = oConn.BeginTransaction();
        oCmd.Transaction = myTrans;
        try
        {
            swClass.AddParameters(oCmd);//宣告參數
            //string yearMonth = IiSheet.GetRow(5).Cells[2].ToString();//value:201501
            for (int i = 0; i < rowcount; i++)
            {
                //try
                //{
                //    dt.Rows[i][0].ToString().Trim();
                //}
                //catch
                //{
                //    continue;
                //}

                if (dt.Rows[i][0].ToString().Trim() != "" && dt.Rows[i][1].ToString().Trim()!="")
                {
                    string yearMonth = string.Empty;
                    DateTime dtYearMonth;
                    string Ipstr = string.Empty;
                    int downloadCount = 0;
                    string dbtype = string.Empty;
                    //int ViewCount = 0;
                    //int logCount = 0;
                    //int indexCount = 0;
                    string documentTitle = string.Empty;
                    //string documentChapter = string.Empty;
                    //string empno = string.Empty;
                    //string email = string.Empty;

                    //string isbn = string.Empty;
                    //string issn = string.Empty;
                    //string main_org = string.Empty;
                    //string main_orgcname = string.Empty;
                    //string main_deptcd = string.Empty;
                    //string main_deptcdcname = string.Empty;
                    try
                    {
                        dbtype = "工程";
                        yearMonth = dt.Rows[i][1].ToString();
                        //dtYearMonth = DateTime.ParseExact(yearMonth, "yyyyMMdd", CultureInfo.InvariantCulture);
                        dtYearMonth = Convert.ToDateTime(yearMonth);
                        //ViewCount = dt.Rows[i][2].ToString().Trim() == "" ? 0 : Convert.ToInt32(dt.Rows[i][2].ToString().Trim());
                        downloadCount = 1;
                        //logCount = dt.Rows[i][4].ToString().Trim() == "" ? 0 : Convert.ToInt32(dt.Rows[i][4].ToString().Trim());
                        //indexCount = dt.Rows[i][5].ToString().Trim() == "" ? 0 : Convert.ToInt32(dt.Rows[i][5].ToString().Trim());
                        documentTitle = dt.Rows[i][4].ToString();
                        //documentChapter = dt.Rows[i][7].ToString();
                        //isbn = dt.Rows[i][8].ToString();
                        //issn = dt.Rows[i][9].ToString();
                        Ipstr = dt.Rows[i][0].ToString();
                        //email = dt.Rows[i][11].ToString();
                        //empno = dt.Rows[i][12].ToString();
                        //usr = dt.Rows[i][6].ToString();

                    }
                    catch (Exception ex)
                    {
                        continue;
                    }

                    oCmd.Parameters["@main_parentid"].Value = strGUID;
                    oCmd.Parameters["@main_ip"].Value = Ipstr;
                    oCmd.Parameters["@main_dbtype"].Value = dbtype;
                    oCmd.Parameters["@main_downloadC"].Value = downloadCount;
                    //oCmd.Parameters["@main_viewC"].Value = ViewCount;
                    //oCmd.Parameters["@main_logC"].Value = logCount;
                    //oCmd.Parameters["@main_indexC"].Value = indexCount;
                    oCmd.Parameters["@main_datetime"].Value = dtYearMonth;
                    oCmd.Parameters["@main_title"].Value = documentTitle;
                    //oCmd.Parameters["@main_chapter"].Value = documentChapter;
                    //oCmd.Parameters["@main_isbn"].Value = isbn;
                    //oCmd.Parameters["@main_issn"].Value = issn;
                    //oCmd.Parameters["@main_email"].Value = email;
                    //oCmd.Parameters["@main_empno"].Value = empno;
                    //oCmd.Parameters["@sender_parentid"].Value = hidGuid;
                    oCmd.Parameters["@main_createdate"].Value = DateTime.Now;

                    oCmd.CommandText = sb.ToString();
                    swClass.RemoveExtraParameters(oCmd, c);//移除不需要用到的參數
                    oCmd.ExecuteNonQuery();
                    c = false;
                    realImportNum++;
                }
                else
                {
                    continue;
                }
            }
            myTrans.Commit();
        }
        catch (Exception ex)
        {
            myTrans.Rollback();
            throw new Exception(ex.Message);
        }
        finally
        {
            oCmd.Connection.Close();
            oConn.Close();

            //if (realImportNum != 0)
            //{
            //    swClass sw = new swClass();
            //    sw.saveLog(realImportNum, strGUID, dbtype);//做通用log
            //}
        }

        return realImportNum;
    }
}