﻿using System;
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
/// s_ACM 的摘要描述
/// </summary>
public class s_ACM
{
	public s_ACM()
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
        sb.Append(@"main_viewC,main_datetime,");
        sb.Append(@"main_createdate");
        //sb.Append(@"");
        //sb.Append(@"");
        sb.Append(@" ) ");
        sb.Append(@" VALUES (");
        sb.Append(@"@main_parentid,@main_ip,@main_dbtype,");
        sb.Append(@"@main_viewC,@main_datetime,");
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
            int howManyYears = IiSheet.GetRow(7).Cells.Count - 2;//月份數會不確定,所以需要先判斷CELL數,再去扣掉多餘的

            for (int i = 9; i < rowcount; i++)
            {
                try
                {
                    IiSheet.GetRow(i).Cells[0].ToString().Trim();
                }
                catch
                {
                    continue;
                }

                for (int j = 2; j < howManyYears + 2; j++)
                {
                    string yearMonth = string.Empty;
                    DateTime dtYearMonth;
                    string Ipstr = string.Empty;
                    int viewCount = 0;

                    if (IiSheet.GetRow(i).Cells[0].ToString().Trim() != "")
                    {
                        try
                        {
                            yearMonth = IiSheet.GetRow(7).Cells[j].ToString();
                            dtYearMonth = Convert.ToDateTime(yearMonth);
                            Ipstr = IiSheet.GetRow(i).Cells[0].ToString();
                            viewCount = Convert.ToInt32(IiSheet.GetRow(i).Cells[j].ToString().Trim());

                            if (viewCount != 0)
                            {
                                oCmd.Parameters["@main_parentid"].Value = strGUID;
                                oCmd.Parameters["@main_ip"].Value = Ipstr;
                                oCmd.Parameters["@main_dbtype"].Value = "ACM";
                                oCmd.Parameters["@main_viewC"].Value = viewCount;
                                oCmd.Parameters["@main_datetime"].Value = dtYearMonth;
                                oCmd.Parameters["@main_createdate"].Value = DateTime.Now;

                                oCmd.CommandText = sb.ToString();
                                swClass.RemoveExtraParameters(oCmd, c);//移除不需要用到的參數
                                oCmd.ExecuteNonQuery();
                                c = false;
                                realImportNum++;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
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