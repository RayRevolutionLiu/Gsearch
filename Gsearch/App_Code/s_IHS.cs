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
/// s_IHS 的摘要描述
/// </summary>
public class s_IHS
{
	public s_IHS()
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
        sb.Append(@"INSERT INTO rs_main ( main_parentid,main_dbtype, ");
        sb.Append(@"main_logC, main_datetime, ");
        sb.Append(@"main_cname, main_org,");
        sb.Append(@"main_orgcname, main_deptcd, main_deptcdcname, main_createdate )");
        sb.Append(@" VALUES (");
        sb.Append(@" @main_parentid,@main_dbtype,@main_logC,");
        sb.Append(@"@main_datetime,@main_cname,@main_org,@main_orgcname,@main_deptcd,@main_deptcdcname,");
        sb.Append(@"@main_createdate ");
        sb.Append(@") ");


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
            for (int i = 1; i < rowcount ; i++)
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
                    string cname = string.Empty;
                    string yearMonth = string.Empty;
                    DateTime dtYearMonth;
                    string org = string.Empty;
                    string deptcd = string.Empty;
                    int logcount = 0;

                    try
                    {
                        cname = IiSheet.GetRow(i).Cells[0].ToString();
                        yearMonth = IiSheet.GetRow(i).Cells[9].ToString();
                        dtYearMonth = Convert.ToDateTime(yearMonth);
                        org = IiSheet.GetRow(i).Cells[1].ToString().Substring(0, 2);
                        deptcd = IiSheet.GetRow(i).Cells[1].ToString().Substring(2, IiSheet.GetRow(i).Cells[1].ToString().Length - 2);
                        logcount = 1;
                    }
                    catch
                    {
                        continue;
                    }

                    oCmd.Parameters["@main_parentid"].Value = strGUID;
                    oCmd.Parameters["@main_dbtype"].Value = "Goldfire";
                    oCmd.Parameters["@main_logC"].Value = logcount;
                    oCmd.Parameters["@main_datetime"].Value = dtYearMonth;
                    oCmd.Parameters["@main_cname"].Value = cname;
                    oCmd.Parameters["@main_org"].Value = org;
                    oCmd.Parameters["@main_orgcname"].Value = OrgCname(org);
                    oCmd.Parameters["@main_deptcd"].Value = deptcd;
                    oCmd.Parameters["@main_deptcdcname"].Value =  deptcdCname(deptcd, org);
                    oCmd.Parameters["@main_createdate"].Value = DateTime.Now;

                    //oCmd.Parameters["@main_ip"].Value = "";
                    //oCmd.Parameters["@main_viewC"].Value = 0;
                    //oCmd.Parameters["@main_downloadC"].Value = 0;
                    //oCmd.Parameters["@main_indexC"].Value = 0;
                    //oCmd.Parameters["@main_title"].Value = "";
                    //oCmd.Parameters["@main_chapter"].Value = "";
                    //oCmd.Parameters["@main_account"].Value = "";
                    //oCmd.Parameters["@main_empno"].Value = "";
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


    public string OrgCname(string org)
    {
        try
        {
            SqlCommand oCmd = new SqlCommand();
            oCmd.Connection = new SqlConnection(AppConfig.DSN);
            StringBuilder sb = new StringBuilder();
            sb.Append(@"SELECT org_abbr_chnm2 FROM common..orgcod WHERE org_orgcd = @org_orgcd ");
            oCmd.CommandText = sb.ToString();
            oCmd.CommandType = CommandType.Text;
            SqlDataAdapter oda = new SqlDataAdapter(oCmd);
            oCmd.Parameters.AddWithValue("@org_orgcd", org);
            DataTable dt = new DataTable();
            oda.Fill(dt);

            return dt.Rows[0][0].ToString().Trim();
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public string deptcdCname(string deptcd, string org)
    {
        try
        {
            SqlCommand oCmd = new SqlCommand();
            oCmd.Connection = new SqlConnection(AppConfig.DSN);
            StringBuilder sb = new StringBuilder();
            sb.Append(@"SELECT dep_deptname FROM common..depcod WHERE dep_deptid = @dep_deptid ");
            oCmd.CommandType = CommandType.Text;
            oCmd.CommandText = sb.ToString();
            SqlDataAdapter oda = new SqlDataAdapter(oCmd);
            oCmd.Parameters.AddWithValue("@dep_deptid", org + deptcd);
            DataTable dt = new DataTable();
            oda.Fill(dt);

            return dt.Rows[0][0].ToString().Trim();
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }
}