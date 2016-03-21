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
/// s_JCR 的摘要描述
/// </summary>
public class s_JCR
{
	public s_JCR()
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
        sb.Append(@"main_parentid,main_ip,main_dbtype,");
        sb.Append(@"main_logC,main_indexC,main_datetime,");
        sb.Append(@"main_createdate");
        //sb.Append(@"");
        //sb.Append(@"");
        sb.Append(@" ) ");
        sb.Append(@" VALUES (");
        sb.Append(@"@main_parentid,@main_ip,@main_dbtype,");
        sb.Append(@"@main_logC,@main_indexC,@main_datetime,");
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

            string yearMonth = string.Empty;
            string Ipstr = string.Empty;
            int logCount = 1;
            int indexCount = 1;

            for (int i = 4; i < rowcount; i++)
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
                    try
                    {
                        Convert.ToDateTime(IiSheet.GetRow(i).Cells[0].ToString());
                        //上一行是為了判斷第一個是否為日期,不是日期就不要去覆蓋掉yearMonth的值,去跑error contiune;
                        yearMonth = IiSheet.GetRow(i).Cells[0].ToString();
                    }
                    catch
                    {
                        continue;
                    }
                }
                else
                {
                    try
                    {
                        Ipstr = IiSheet.GetRow(i).Cells[1].ToString().Trim();
                        logCount = Convert.ToInt32(IiSheet.GetRow(i).Cells[2].ToString().Trim());
                        indexCount = Convert.ToInt32(IiSheet.GetRow(i).Cells[3].ToString().Trim());
                    }
                    catch
                    {
                        continue;
                    }

                    oCmd.Parameters["@main_parentid"].Value = strGUID;
                    oCmd.Parameters["@main_ip"].Value = Ipstr;
                    oCmd.Parameters["@main_dbtype"].Value = "JCR";
                    oCmd.Parameters["@main_logC"].Value = logCount;
                    oCmd.Parameters["@main_indexC"].Value = indexCount;
                    oCmd.Parameters["@main_datetime"].Value = Convert.ToDateTime(yearMonth);
                    //oCmd.Parameters["@main_title"].Value = documentTitle;
                    //oCmd.Parameters["@main_account"].Value = usr;
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