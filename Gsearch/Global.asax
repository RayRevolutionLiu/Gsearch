<%@ Application Language="C#" %>
<script runat="server">

    void Application_Start(object sender, EventArgs e) 
    {
        // 在應用程式啟動時執行的程式碼
        //把single sign on 之後的工號抓到
        //string getAD = GetSSOAttribute("SM_USER");

        //AccountInfo accInfo = new Account().ExecLogon(getAD);
        ////如果accinfo不等於空值
        //if (accInfo != null)
        //{
        //    //將該物件accinfo傳給Session["AccountInfo"]保存
        //    Session["pwerRowData"] = accInfo;

        //}
        //else
        //{
        //    Application["ErrorMsg"] = "您沒有權限";
        //    Response.Redirect("~/errorPage.aspx");
        //}
    }

    public string GetSSOAttribute(string AttrName)
    {

        if (System.Configuration.ConfigurationManager.AppSettings["DEBUG"].ToUpper() == "TRUE")
        {
            return System.Configuration.ConfigurationManager.AppSettings["DEBUG.ACCOUNT"];
        }

        string AllHttpAttrs, FullAttrName, Result;
        int AttrLocation;
        AllHttpAttrs = Request.ServerVariables["ALL_HTTP"];
        FullAttrName = "HTTP_" + AttrName.ToUpper();
        AttrLocation = AllHttpAttrs.IndexOf(FullAttrName + ":");
        if (AttrLocation > 0)
        {
            Result = AllHttpAttrs.Substring(AttrLocation + FullAttrName.Length + 1);
            AttrLocation = Result.IndexOf("\n");
            if (AttrLocation <= 0) AttrLocation = Result.Length + 1;
            //return Result.Substring(0, AttrLocation - 1);
            return HttpContext.Current.Request.ServerVariables["HTTP_SM_USER"];
        }

        return "";
    }
    
    void Application_End(object sender, EventArgs e) 
    {
        //  在應用程式關閉時執行的程式碼

    }
        
    void Application_Error(object sender, EventArgs e) 
    { 
        // 在發生未處理的錯誤時執行的程式碼
        bool IsExecAjax = false;
        Regex regexFormat = new Regex(@"exec\S*\.aspx", RegexOptions.IgnoreCase);
        IsExecAjax = (regexFormat.Match(Request.Path).Success == true) ? true : false;

        Exception ex = Server.GetLastError().GetBaseException();
        string Message = string.Format("訊息：{0}", ex.GetBaseException().Message);


        /*===========explain：show message*/
        Server.ClearError();

        if (IsExecAjax)
        {
            Response.Write(Message);
        }
        else
        {
            //if (Context.Response.StatusCode == 404)
            //{

            //}
            Application["ErrorMsg"] = Message;
            Response.Redirect("~/errorPage.aspx");

        }
    }

    void Session_Start(object sender, EventArgs e) 
    {
        // 在新的工作階段啟動時執行的程式碼

    }

    void Session_End(object sender, EventArgs e) 
    {
        // 在工作階段結束時執行的程式碼
        // 注意: 只有在  Web.config 檔案中將 sessionstate 模式設定為 InProc 時，
        // 才會引起 Session_End 事件。如果將 session 模式設定為 StateServer 
        // 或 SQLServer，則不會引起該事件。

    }
       
</script>
