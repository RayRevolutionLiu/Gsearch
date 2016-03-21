<%@ WebHandler Language="C#" Class="filedownload" %>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;
using System.Data;
using System.IO;

public class filedownload : IHttpHandler, System.Web.SessionState.IRequiresSessionState
{
    
    public void ProcessRequest (HttpContext context) {
        try
        {
            string selected = context.Request.Form["selectedIn"].ToString().Trim();
            string selectedOrg = context.Request.Form["selectedOrgIn"].ToString().Trim();
            string main_title = context.Request.Form["main_titleIn"].ToString().Trim();
            string main_chapter = context.Request.Form["main_chapterIn"].ToString().Trim();
            string main_empno = context.Request.Form["main_empnoIn"].ToString().Trim();
            string main_cname = context.Request.Form["main_cnameIn"].ToString().Trim();
            string main_deptcd = context.Request.Form["main_deptcdIn"].ToString().Trim();
            string main_startdate = context.Request.Form["main_startdateIn"].ToString().Trim();
            string main_enddate = context.Request.Form["main_enddateIn"].ToString().Trim();
            string sortcolumn = context.Request.Form["sortcolumnIn"].ToString().Trim();
            string orderBy = context.Request.Form["orderByIn"].ToString().Trim();
            string Export = context.Request.Form["ExportIn"].ToString().Trim();

            SqlCommand oCmd = new SqlCommand();
            oCmd.Connection = new SqlConnection(AppConfig.DSN);
            StringBuilder sb = new StringBuilder();
            sb.Append(@"GenTable");
            //sb.Append(@"");
            oCmd.CommandText = sb.ToString();
            oCmd.CommandType = CommandType.StoredProcedure;
            oCmd.Parameters.AddWithValue("@selected", selected);
            oCmd.Parameters.AddWithValue("@selectedOrg", selectedOrg);
            oCmd.Parameters.AddWithValue("@main_title", main_title);
            oCmd.Parameters.AddWithValue("@main_chapter", main_chapter);
            oCmd.Parameters.AddWithValue("@main_empno", main_empno);
            oCmd.Parameters.AddWithValue("@main_cname", main_cname);
            oCmd.Parameters.AddWithValue("@main_deptcd", main_deptcd);
            oCmd.Parameters.AddWithValue("@main_startdate", main_startdate);
            oCmd.Parameters.AddWithValue("@main_enddate", main_enddate);

            oCmd.Parameters.AddWithValue("@START", string.Empty);
            oCmd.Parameters.AddWithValue("@END", string.Empty);

            oCmd.Parameters.AddWithValue("@sortcolumn", sortcolumn);
            oCmd.Parameters.AddWithValue("@orderBy", orderBy);
            oCmd.Parameters.AddWithValue("@Export", Export);
            SqlDataAdapter oda = new SqlDataAdapter(oCmd);
            DataTable dt = new DataTable();
            oda.Fill(dt);


            //丟入匯出function
            ExportExcel ex = new ExportExcel();
            ex.dt = dt;
            string[] header = { "資料庫名稱", "IP", "日期", "單位", "部門名稱", "部門代碼", "組別", "姓名", "工號", "職務", "職級", "年資", "瀏覽次數", "下載次數", "登入次數", "檢索次數", "信箱", "刊名", "篇名" };
            ex.header = header;
            ex.CreateExcel(null);
            /*底下是用新的工具匯出Excel,不過一樣會跳出錯誤,不過Excel的內容卻正確 */
            //ex.CreatExcelOpenXml();
            
            //if (PassStream(context.Response.OutputStream, ex.CreateExcel(), context))
            //{
            //    context.Response.Write("<script type='text/JavaScript'>parent.ajaxAlertFun('success'); </script>");
            //}
            ex.CreateExcel(null).Close();
            ex.CreateExcel(null).Dispose();
            context.Session["downloadCSV"] = true;
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }


    /// <summary>
    /// 以分批的方式傳送串流間的資料, 以減少對記憶體的需求
    /// 預設緩衝區(buffer)=8192(byte)，可經由設定大小，來控制下載速度及對記憶體的需求
    /// </summary>
    /// <param name="outStream">資料輸出串流</param>
    /// <param name="inStream">資料輸入串流</param>
    bool PassStream(Stream outStream, Stream inStream, HttpContext context)
    {
        int bufflen = 8192;
        int len = 0;
        byte[] buff = new byte[bufflen];
        len = inStream.Read(buff, 0, bufflen);
        while (len > 0)
        {
            //檢查用戶是否能在連線,  如果有才繼續傳送封包
            if (context.Response.IsClientConnected)
            {

                outStream.Write(buff, 0, len);
                len = inStream.Read(buff, 0, bufflen);

            }
            else
            {
                len = 0;
            }

        }
        outStream.Flush();
        return true;
    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}