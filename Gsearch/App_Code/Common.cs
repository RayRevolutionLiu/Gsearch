using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Web;
using System.Data;
using System.Text;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using System.Globalization;

/// <summary>
/// Common 的摘要描述
/// </summary>
public class Common
{
	public Common()
	{
		//
		// TODO: 在這裡新增建構函式邏輯
		//
	}

    //新增百分比到Datatable內
    public DataTable AddPercentageToDataTable(DataTable dt)
    {
        if (dt.Rows.Count > 0)
        {
            dt.Columns.Add("main_percentage", System.Type.GetType("System.String"));

            int TotlaCount = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                TotlaCount += Convert.ToInt32(dt.Rows[i]["total"].ToString().Trim());
            }

            for (int j = 0; j < dt.Rows.Count; j++)
            {
                decimal result = Convert.ToInt32(dt.Rows[j]["total"].ToString()) == 0 ? 0 : Math.Round((decimal)Convert.ToInt32(dt.Rows[j]["total"].ToString()) / TotlaCount, 2);

                dt.Rows[j]["main_percentage"] = String.Format(CultureInfo.InvariantCulture, "{0:#0.##%}", result);
            }

            dt.Columns["main_percentage"].SetOrdinal(2);
            return dt;
        }
        else
        {
            return dt;
        }
    }
}

/// <summary>
/// MbrAccount 的摘要描述。
/// </summary>
public class Account
{
    string conn = ConfigurationManager.AppSettings["DSN.charity"].ToString();
    /// <summary>
    /// GetAccInfo
    /// </summary>
    /// <returns></returns>
    public static AccountInfo GetAccInfo()
    {
        return (AccountInfo)HttpContext.Current.Session["pwerRowData"];
    }


    /// <summary>
    /// 進行帳號登入檢查，通過後會取得登入者資料並放入session:ds_CurrentUser
    /// </summary>
    /// <param name="inputID"></param>
    /// <returns></returns>
    public AccountInfo ExecLogon(string getAD)
    {
        DataSet ds = new DataSet();
        SqlDataAdapter oda = new SqlDataAdapter();
        SqlConnection thisConnection = new SqlConnection(conn);
        SqlCommand thisCommand = thisConnection.CreateCommand();
        StringBuilder show_value = new StringBuilder();
        try
        {
            thisConnection.Open();
            show_value.Append(@"select * from common..comper where com_empno=@getAD ");
            thisCommand.Parameters.AddWithValue("@getAD", getAD);
            thisCommand.CommandType = CommandType.Text;
            thisCommand.CommandText = show_value.ToString();
            oda.SelectCommand = thisCommand;
            oda.Fill(ds);
        }
        finally
        {
            oda.Dispose();
            thisConnection.Close();
            thisConnection.Dispose();
            thisCommand.Dispose();

        }

        if (ds.Tables[0].Rows.Count > 0)
        {
            return new AccountInfo(ds);
        }
        else
        {
            return null;
        }
    }

}

/*-------------------------------------------------------------------------------------------------------------------------*/

/// <summary>
/// MbrInfo 的摘要描述。
/// </summary>
public class AccountInfo
{
    /// <summary>
    /// 工號
    /// </summary>
    public readonly string com_empno = "";
    /// <summary>         
    /// 姓名              
    /// </summary>        
    public readonly string com_cname = "";
    /// <summary>          
    /// 電話               
    /// </summary>         
    public readonly string com_telext = "";
    /// <summary>          
    /// 身分類別           
    /// </summary>         
    public readonly string com_atype = "";
    /// <summary>          
    /// 單位               
    /// </summary>         
    public readonly string com_orgcd = "";
    /// <summary>          
    /// 部門               
    /// </summary>         
    public readonly string com_deptcd = "";


    public AccountInfo(DataSet ds)
    {
        if (ds.Tables[0].Rows.Count > 0)
        {
            DataRow dr = ds.Tables[0].Rows[0];

            com_empno = dr["com_empno"].ToString();
            com_cname = dr["com_cname"].ToString();
            com_telext = dr["com_telext"].ToString();
            com_orgcd = dr["com_orgcd"].ToString();
            com_deptcd = dr["com_deptcd"].ToString();
            //if (mbrAccid.ToLower() == "guest")
            //{
            //    isGuestRole = true;
            //}
            //else if (mbrAccid.ToLower() != "guest")
            //{
            //    if (mbrTypeStatus == "1")
            //    {
            //        isApplyMbrRole = true;
            //    }
            //    else if (mbrTypeStatus == "2")
            //    {
            //        isApproveMbrRole = true;
            //    }


            //    if (mbrTypeStatus == "2" && mbrTypeAdmin == "1")
            //    {
            //        isAdminRole = true;
            //    }


            //    //if (mbrTypeClass == "mbr4")
            //    //{
            //    //    mbrTypeClass2 = "MEMBER04";
            //    //    isMbrClassOfService = true;
            //    //}
            //    //else if (mbrTypeClass == "mbr5")
            //    //{
            //    //    mbrTypeClass2 = "MEMBER05";
            //    //    isMbrClassOfPOrT = true;
            //    //    isMbrClassOfTeam = true;
            //    //}
            //    //else if (mbrTypeClass == "mbr6")
            //    //{
            //    //    mbrTypeClass2 = "MEMBER06";
            //    //    isMbrClassOfPOrT = true;
            //    //    isMbrClassOfPerson = true;
            //}
        }
    }
}

#region 加解密
/// <summary>
/// encryptquerystring 加密QueryString decryptquerystring 解密
/// encryptpassword 加密密碼 decryptpassword 解密
/// 用法 string aaa = encryptquerystring(要加密的字串);
/// </summary>
public class security
{
    string _querystringkey = "usagelog"; //url传输参数加密key
    string _passwordkey = "usagelog"; //password加密key

    public security()
    {
        //
        // todo: 在此处添加构造函数逻辑
        //
    }

    /// 
    /// 加密url传输的字符串
    /// 
    /// 
    /// 
    public string encryptquerystring(string querystring)
    {
        return Encrypt(querystring, _querystringkey);
    }

    /// 
    /// 解密url传输的字符串
    /// 
    /// 
    /// 
    public string decryptquerystring(string querystring)
    {
        return decrypt(querystring, _querystringkey);
    }

    /// 
    /// 加密帐号口令
    /// 
    /// 
    /// 
    public string encryptpassword(string password)
    {
        return Encrypt(password, _passwordkey);
    }

    /// 
    /// 解密帐号口令
    /// 
    /// 
    /// 
    public string decryptpassword(string password)
    {
        return decrypt(password, _passwordkey);
    }

    /// 
    /// dec 加密过程
    /// 
    /// 
    /// 
    /// 
    public string Encrypt(string pToEncrypt, string sKey)
    {
        DESCryptoServiceProvider des = new DESCryptoServiceProvider(); //把字符串放到byte數組中 

        byte[] inputByteArray = Encoding.Default.GetBytes(pToEncrypt);
        //byte[] inputByteArray=Encoding.Unicode.GetBytes(pToEncrypt); 

        des.Key = ASCIIEncoding.ASCII.GetBytes(sKey); //建立加密對象的密鑰和偏移量 
        des.IV = ASCIIEncoding.ASCII.GetBytes(sKey); //原文使用ASCIIEncoding.ASCII方法的GetBytes方法 
        MemoryStream ms = new MemoryStream(); //使得輸入密碼必須輸入英文文本 
        CryptoStream cs = new CryptoStream(ms, des.CreateEncryptor(), CryptoStreamMode.Write);

        cs.Write(inputByteArray, 0, inputByteArray.Length);
        cs.FlushFinalBlock();

        StringBuilder ret = new StringBuilder();
        foreach (byte b in ms.ToArray())
        {
            ret.AppendFormat("{0:X2}", b);
        }
        ret.ToString();
        return ret.ToString();
    }

    /// 
    /// dec 解密过程
    /// 
    /// 
    /// 
    /// 
    public string decrypt(string ptodecrypt, string skey)
    {
        DESCryptoServiceProvider des = new DESCryptoServiceProvider();

        byte[] inputbytearray = new byte[ptodecrypt.Length / 2];
        for (int x = 0; x < ptodecrypt.Length / 2; x++)
        {
            int i = (Convert.ToInt32(ptodecrypt.Substring(x * 2, 2), 16));
            inputbytearray[x] = (byte)i;
        }

        des.Key = ASCIIEncoding.ASCII.GetBytes(skey); //建立加密对象的密钥和偏移量，此值重要，不能修改 
        des.IV = ASCIIEncoding.ASCII.GetBytes(skey);
        MemoryStream ms = new MemoryStream();
        CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(), CryptoStreamMode.Write);

        cs.Write(inputbytearray, 0, inputbytearray.Length);
        cs.FlushFinalBlock();

        StringBuilder ret = new StringBuilder(); //建立stringbuild对象，createdecrypt使用的是流对象，必须把解密后的文本变成流对象 

        return System.Text.Encoding.Default.GetString(ms.ToArray());
    }

    /// 
    /// 检查己加密的字符串是否和原文相同
    /// 
    /// 
    /// 
    /// 
    /// 
    public bool validatestring(string enstring, string fostring, int mode)
    {
        switch (mode)
        {
            default:
            case 1:
                if (decrypt(enstring, _querystringkey) == fostring.ToString())
                {
                    return true;
                }
                else
                {
                    return false;
                }
            case 2:
                if (decrypt(enstring, _passwordkey) == fostring.ToString())
                {
                    return true;
                }
                else
                {
                    return false;
                }
        }
    }
}
#endregion