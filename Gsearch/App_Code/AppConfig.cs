using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

/// <summary>
/// AppConfig 的摘要描述
/// </summary>
public class AppConfig
{
    public static string _DSN;

    static AppConfig()
    {
        //
        // TODO: 在這裡新增建構函式邏輯
        //
        _DSN = ReadConfigString("DSN.charity").ToString();
    }

    private static string ReadConfigString(string keyName)
    {
        string retStr = ConfigurationManager.AppSettings[keyName];
        if (retStr == null)
        {
            throw new Exception("Unable to read app settings.Please check appsetting section in Web.config file.(" + keyName + ")");
        }
        return retStr;
    }

    public static string DSN
    {
        get { return AppConfig._DSN; }
    }
}