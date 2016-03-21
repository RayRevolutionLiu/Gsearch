<%@ WebHandler Language="C#" Class="deleteAnnualData" %>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;
using System.Data;

public class deleteAnnualData : IHttpHandler {

    public void ProcessRequest(HttpContext context)
    {
        SqlConnection oConn = new SqlConnection(AppConfig.DSN);
        oConn.Open();
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = oConn; 
        oCmd.CommandType = CommandType.Text;
        StringBuilder sb = new StringBuilder();
        
        try
        {
            string year = context.Request["year"].ToString().Trim();
            string uid = context.Request["uid"].ToString().Trim();
            
            //for SQL injection
            Convert.ToInt32(year);
            if (uid.Trim() != "")
            {
                Guid.Parse(uid);

            }
               
            if (uid == "")
            {
                sb.Append(@"DELETE FROM rs_main WHERE YEAR(main_createdate) = @year ; DELETE FROM rs_statistics WHERE YEAR(statistics_datetime) = @year ");
                oCmd.Parameters.AddWithValue("@year", year);
            }
            else
            {
                sb.Append(@"DELETE FROM rs_main WHERE main_parentid = @uid ; DELETE FROM rs_statistics WHERE statistics_parentid = @uid ");
                oCmd.Parameters.AddWithValue("@uid", uid);
            }
            
            oCmd.CommandText = sb.ToString();
            oCmd.ExecuteNonQuery();

        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
        finally
        {
            oCmd.Connection.Close();
            oConn.Close();
        }
    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}