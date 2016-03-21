<%@ WebHandler Language="C#" Class="importexcel" %>

using System;
using System.Web;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Data.SqlClient;
using System.Text;
using System.Data;

public class importexcel : IHttpHandler {
    
    public void ProcessRequest (HttpContext context) {
        try
        {
            //測試用
            //context.Response.Write("<script type='text/JavaScript'>parent.ajaxAlertFun('error_123 4340 93284 0932ru20 u20fju42 0fju093f j20jf02fj0 2jf20fj kewpofkw ewofkewofwepok pow kpoewkf poewf kpewofk powefk fpowk fe kwp kfjoifjewf jwoifjoiew oijewfj iewwpoiej ijfewj oifew'); </script>");
            //return;
            
            HttpFileCollection uploadFiles = context.Request.Files;//檔案集合
            string textType = context.Request.Form["textType"].ToString().Trim();
            string YearVal = Convert.ToInt32(context.Request.Form["YearVal"].ToString().Trim()).ToString();//如果值不是數字才可以被catch出來
            string MonthVal = Convert.ToInt32(context.Request.Form["MonthVal"].ToString().Trim()).ToString().Length == 1 ? "0" + context.Request.Form["MonthVal"].ToString().Trim() : context.Request.Form["MonthVal"].ToString().Trim();//如果值不是數字才可以被catch出來
            
            if (uploadFiles.Count > 1)
            {
                throw new Exception("一次請選擇一個檔案");
            }
            
            //如果有檔案
            if (uploadFiles.Count > 0)
            {
                HttpPostedFile aFile = uploadFiles[0];
                string extension = (System.IO.Path.GetExtension(aFile.FileName) == "") ? "" : System.IO.Path.GetExtension(aFile.FileName);
                if (extension != ".xls" && extension != ".xlsx")
                {
                    throw new Exception("請選擇xls或xlsx檔案上傳");
                }

                //phase 2 如果此類別再當月已經被上傳過一次了,就必須要讓警告
                if (DidyouUploadthisMonth(textType, YearVal, MonthVal) && textType.ToString().Trim() != "通用")
                {
                    context.Response.Write("<script type='text/JavaScript'>parent.ajaxAlertFun('error_此月份已上傳過此類型檔案'); </script>");
                }
                else if (DateTime.Now.Day == 1)
                {
                    context.Response.Write("<script type='text/JavaScript'>parent.ajaxAlertFun('error_今日為一號,請勿於今日上傳任何資料'); </script>");
                }
                else
                {

                    IWorkbook workbook;// = new HSSFWorkbook();//创建Workbook对象
                    if (extension == ".xls")
                    {
                        workbook = new HSSFWorkbook(aFile.InputStream);
                    }
                    else
                    {
                        workbook = new XSSFWorkbook(aFile.InputStream);
                    }

                    ISheet sheet = workbook.GetSheetAt(0);
                    swClass sw = new swClass();
                    sw.IiSheet = sheet;
                    sw.TextType = textType;
                    sw.Year = YearVal;
                    sw.Month = MonthVal;
                    int cou = sw.Execute();

                    if (cou == 0)
                    {
                        context.Response.Write("<script type='text/JavaScript'>parent.ajaxAlertFun('error_上傳失敗，上傳格式錯誤'); </script>");
                    }
                    else
                    {
                        context.Response.Write("<script type='text/JavaScript'>parent.ajaxAlertFun('正常'); </script>");
                    }
                }
            }      
        }
        catch (Exception ex)
        {
            context.Response.Write("<script type='text/JavaScript'>parent.ajaxAlertFun('error_" + ex.Message.Replace("\r\n","") + "'); </script>");
        }
            
    }

    public bool DidyouUploadthisMonth(string dbtype,string Year,string Month)
    {
        string yearMonth = Year + Month + "01";
        DateTime dTime = DateTime.ParseExact(yearMonth, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
        
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(AppConfig.DSN);
        StringBuilder sb = new StringBuilder();
        sb.Append(@"select * FROM rs_statistics WHERE substring(CONVERT(varchar(12) ,statistics_datetime, 112 ),1,6) = substring(CONVERT(varchar(12) ,@statistics_datetime, 112 ),1,6) AND statistics_dbtype = @statistics_dbtype ");
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        oCmd.Parameters.AddWithValue("@statistics_dbtype", dbtype);
        oCmd.Parameters.AddWithValue("@statistics_datetime", dTime);
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();
        oda.Fill(dt);

        if (dt.Rows.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}