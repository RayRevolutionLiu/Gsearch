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
/// s_WangFang 的摘要描述
/// </summary>
public class s_WangFang
{
	public s_WangFang()
	{
		//
		// TODO: 在這裡新增建構函式邏輯
		//
	}

    public int Execute(ISheet IiSheet, string strGUID)
    {
        int rowcount = IiSheet.PhysicalNumberOfRows;
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
            for (int i = 1; i < rowcount; i++)
            {
                try
                {
                    IiSheet.GetRow(i).Cells[5].ToString().Trim();
                }
                catch
                {
                    continue;
                }

                if (IiSheet.GetRow(i).Cells[5].ToString().Trim() != "")
                {
                    string yearMonth = string.Empty;
                    DateTime dtYearMonth;
                    string Ipstr = string.Empty;
                    int downloadCount = 0;
                    string documentTitle = string.Empty;

                    try
                    {
                        //此excel怪怪的,subject此column會被跳過,所以後面的column要順著下去
                        yearMonth = IiSheet.GetRow(i).Cells[4].ToString();
                        //dtYearMonth = DateTime.ParseExact(yearMonth, "yyyyMMdd", CultureInfo.InvariantCulture);
                        dtYearMonth = Convert.ToDateTime(yearMonth);
                        Ipstr = IiSheet.GetRow(i).Cells[5].ToString();
                        documentTitle = IiSheet.GetRow(i).Cells[2].ToString();
                        //usr = IiSheet.GetRow(i).Cells[6].ToString();
                        downloadCount = 1;//沒有次數 就一筆算一次
                    }
                    catch (Exception ex)
                    {
                        continue;
                    }

                    oCmd.Parameters["@main_parentid"].Value = strGUID;
                    oCmd.Parameters["@main_ip"].Value = Ipstr;
                    oCmd.Parameters["@main_dbtype"].Value = "萬方";
                    oCmd.Parameters["@main_downloadC"].Value = downloadCount;
                    oCmd.Parameters["@main_datetime"].Value = dtYearMonth;
                    oCmd.Parameters["@main_title"].Value = documentTitle;
                    //oCmd.Parameters["@main_account"].Value = usr;
                    //oCmd.Parameters["@sender_parentid"].Value = hidGuid;
                    //oCmd.Parameters["@sender_parentid"].Value = hidGuid;
                    //oCmd.Parameters["@sender_parentid"].Value = hidGuid;
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
            //context.Response.ContentType = "text/html";
            //context.Response.Write("<script type='text/JavaScript'>parent.feedbackFun('" + hidGuid + "');</script>");
        }

        return realImportNum;
    }
}