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
/// s_TI 的摘要描述
/// </summary>
public class s_TI
{
	public s_TI()
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
        sb.Append(@"main_parentid,main_dbtype,main_cname,main_org,main_deptcd,");
        sb.Append(@"main_viewC,main_datetime,");
        sb.Append(@"main_createdate");
        //sb.Append(@"");
        //sb.Append(@"");
        sb.Append(@" ) ");
        sb.Append(@" VALUES (");
        sb.Append(@"@main_parentid,@main_dbtype,@main_cname,@main_org,@main_deptcd,");
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
            //string yearMonth = IiSheet.GetRow(5).Cells[2].ToString();//value:201501
            for (int i = 1; i < rowcount; i++)
            {
                try
                {
                    IiSheet.GetRow(i).Cells[0].ToString().Trim();
                    IiSheet.GetRow(i).Cells[1].ToString().Trim();
                }
                catch
                {
                    continue;
                }

                if (IiSheet.GetRow(i).Cells[0].ToString().Trim() != "")
                {
                    string yearMonth = IiSheet.GetRow(i).Cells[4].ToString();
                    string cname = IiSheet.GetRow(i).Cells[0].ToString();
                    //string deptid = IiSheet.GetRow(i).Cells[1].ToString().Trim();
                    //string[] sArray = ReturnEmpno(cname, deptid);
                    string org = IiSheet.GetRow(i).Cells[1].ToString().Substring(0, 2);
                    string deptcd = IiSheet.GetRow(i).Cells[1].ToString().Substring(2, IiSheet.GetRow(i).Cells[1].ToString().Length - 2);
                    //string main_empno = sArray[0].ToString().Trim();
                    //string emailadd = sArray[1].ToString().Trim();
                    //string org = sArray[2].ToString().Trim();
                    //string orgcname = sArray[4].ToString().Trim();
                    //string deptcd = sArray[3].ToString().Trim();
                    //string deptcname = sArray[5].ToString().Trim();

                    DateTime dtYearMonth = DateTime.ParseExact(yearMonth, "yyyyMMdd", CultureInfo.InvariantCulture);
                    //DateTime dtYearMonth = Convert.ToDateTime(yearMonth);
                    //string Ipstr = string.Empty;
                    //string documentTitle = string.Empty;
                    //string usr = string.Empty;
                    int viewCount = 1;

                    try
                    {
                        //Ipstr = System.Text.RegularExpressions.Regex.Replace(IiSheet.GetRow(i).Cells[4].ToString(), "0*([0-9]+)", "${1}");
                        
                        IiSheet.GetRow(i).Cells[4].ToString();
                        //usr = IiSheet.GetRow(i).Cells[0].ToString();
                    }
                    catch
                    {
                        continue;
                    }

                    oCmd.Parameters["@main_parentid"].Value = strGUID;
                    //oCmd.Parameters["@main_ip"].Value = Ipstr;
                    oCmd.Parameters["@main_dbtype"].Value = "TI";
                    oCmd.Parameters["@main_viewC"].Value = viewCount;
                    oCmd.Parameters["@main_datetime"].Value = dtYearMonth;
                    oCmd.Parameters["@main_cname"].Value = cname;
                    //oCmd.Parameters["@main_empno"].Value = main_empno;
                    //oCmd.Parameters["@main_email"].Value = emailadd;
                    oCmd.Parameters["@main_org"].Value = org;
                    //oCmd.Parameters["@main_orgcname"].Value = orgcname;
                    oCmd.Parameters["@main_deptcd"].Value = deptcd;
                    //oCmd.Parameters["@main_deptcdcname"].Value = deptcname;
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
                    break;
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

    //public string[] ReturnEmpno(string cname,string deptid)
    //{
    //    DataSet ds = new DataSet();
    //    SqlDataAdapter oda = new SqlDataAdapter();
    //    SqlConnection thisConnection = new SqlConnection(AppConfig.DSN);
    //    SqlCommand thisCommand = thisConnection.CreateCommand();
    //    StringBuilder show_value = new StringBuilder();
    //    try
    //    {
    //        thisConnection.Open();
    //        show_value.Append(@"select com_empno,com_mailadd,com_orgcd,com_deptcd,");
    //        show_value.Append(@"(select org_abbr_chnm2 from common..orgcod WHERE com_orgcd = org_orgcd) AS orgcname,");
    //        show_value.Append(@"(select dep_deptname from common..depcod WHERE com_deptid = dep_deptid) AS dep_deptname ");
    //        show_value.Append(@"FROM common..comper where com_cname=@com_cname AND com_deptid=@com_deptid ");
    //        //show_value.Append(@"");
    //        //show_value.Append(@"");
    //        thisCommand.Parameters.AddWithValue("@com_cname", cname);
    //        thisCommand.Parameters.AddWithValue("@com_deptid", deptid);
    //        thisCommand.CommandType = CommandType.Text;
    //        thisCommand.CommandText = show_value.ToString();
    //        oda.SelectCommand = thisCommand;
    //        oda.Fill(ds);
    //    }
    //    finally
    //    {
    //        oda.Dispose();
    //        thisConnection.Close();
    //        thisConnection.Dispose();
    //        thisCommand.Dispose();

    //    }

    //    if (ds.Tables[0].Rows.Count > 0)
    //    {
    //        //理論上不太可能會超過一筆,但是若真的超過一筆就抓第一個人
    //        string[] sArray = new string[ds.Tables[0].Columns.Count];

    //        for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
    //        {
    //            sArray[i] = ds.Tables[0].Rows[0][i].ToString().Trim();
    //        }
    //        return sArray;
    //    }
    //    else
    //    {
    //        return null;
    //    }
    //}
}