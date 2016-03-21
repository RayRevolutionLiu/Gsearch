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

/// <summary>
/// s_nature 的摘要描述
/// </summary>
public class s_nature
{
	public s_nature()
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
        sb.Append(@"main_parentid,main_ip,main_dbtype,main_title,main_chapter,");
        sb.Append(@"main_downloadC,");
        sb.Append(@"main_createdate");
        //sb.Append(@"");
        //sb.Append(@"");
        sb.Append(@" ) ");
        sb.Append(@" VALUES (");
        sb.Append(@"@main_parentid,@main_ip,@main_dbtype,@main_title,@main_chapter,");
        sb.Append(@"@main_downloadC,");
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
            for (int i = 8; i < rowcount ; i++)
            {
                try
                {
                    IiSheet.GetRow(i).Cells[0].ToString().Trim();
                }
                catch
                {
                    continue;
                }

                if (IiSheet.GetRow(i).Cells[0].ToString().Trim() != "")
                {
                    //string yearMonth = string.Empty;
                    //DateTime dtYearMonth;
                    string Ipstr = string.Empty;
                    string documentTitle = string.Empty;
                    string chapter = string.Empty;
                    int downloadCount = 0;

                    try
                    {
                        //yearMonth = IiSheet.GetRow(i).Cells[0].ToString();
                        //dtYearMonth = DateTime.ParseExact(yearMonth, "yyyyMMdd", CultureInfo.InvariantCulture);
                        //dtYearMonth = Convert.ToDateTime(yearMonth);
                        Ipstr = IiSheet.GetRow(i).Cells[5].ToString();
                        documentTitle = IiSheet.GetRow(i).Cells[0].ToString();
                        chapter = IiSheet.GetRow(i).Cells[2].ToString();
                        downloadCount = Convert.ToInt32(IiSheet.GetRow(i).Cells[6].ToString());
                    }
                    catch
                    {
                        continue;
                    }

                    oCmd.Parameters["@main_parentid"].Value = strGUID;
                    oCmd.Parameters["@main_ip"].Value = Ipstr;
                    oCmd.Parameters["@main_dbtype"].Value = "nature";
                    oCmd.Parameters["@main_downloadC"].Value = downloadCount;
                    //oCmd.Parameters["@main_datetime"].Value = dtYearMonth;
                    oCmd.Parameters["@main_title"].Value = documentTitle;
                    oCmd.Parameters["@main_chapter"].Value = chapter;
                    //oCmd.Parameters["@sender_parentid"].Value = hidGuid;
                    //oCmd.Parameters["@sender_parentid"].Value = hidGuid;
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