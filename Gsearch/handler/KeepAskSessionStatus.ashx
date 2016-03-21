<%@ WebHandler Language="C#" Class="KeepAskSessionStatus" %>

using System;
using System.Web;

public class KeepAskSessionStatus : IHttpHandler, System.Web.SessionState.IRequiresSessionState
{
    
    public void ProcessRequest (HttpContext context) {

        if (context.Session["downloadCSV"] != null)
        {
            context.Session["downloadCSV"] = null;
            context.Response.Write("success");
        }
    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}