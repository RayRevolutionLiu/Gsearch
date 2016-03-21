using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using System.Globalization;

/// <summary>
/// swClass 的摘要描述
/// </summary>
public class swClass
{
	public swClass()
	{
		//
		// TODO: 在這裡新增建構函式邏輯
		//
	}

    private string _TextType;
    private ISheet _IiSheet;
    private string _Year;
    private string _Month;

    public string TextType
    {
        get { return this._TextType; }
        set { this._TextType = value; }
    }

    public string Year
    {
        get { return this._Year; }
        set { this._Year = value; }
    }

    public string Month
    {
        get { return this._Month; }
        set { this._Month = value; }
    }

    public ISheet IiSheet
    {
        get { return this._IiSheet; }
        set { this._IiSheet = value; }
    }

    public int Execute()
    {
        string strGUID = Guid.NewGuid().ToString();
        int realImportNum = 0;
        Excel _excel = new Excel();
        DataTable dt = new DataTable();

        switch (TextType)
        {
            case "ASTM":
                s_ASTM ASTM = new s_ASTM();
                realImportNum = ASTM.Execute(IiSheet, strGUID);
                break;
            case "Digitimes":
                s_Digitimes Digi = new s_Digitimes();
                realImportNum = Digi.Execute(IiSheet, strGUID);
                break;
            case "EV":
                s_EV EV = new s_EV();
                realImportNum = EV.Execute(IiSheet, strGUID);
                break;
            case "Goldfire":
                s_IHS IHS = new s_IHS();
                realImportNum = IHS.Execute(IiSheet, strGUID);
                break;
            case "JCR":
                s_JCR JCR = new s_JCR();
                realImportNum = JCR.Execute(IiSheet, strGUID);
                break;
            case "nature":
                s_nature nature = new s_nature();
                realImportNum = nature.Execute(IiSheet, strGUID);
                break;
            case "Oxford":
                s_Oxford Oxford = new s_Oxford();
                realImportNum = Oxford.Execute(IiSheet, strGUID);
                break;
            case "Science":
                s_Science Science = new s_Science();
                realImportNum = Science.Execute(IiSheet, strGUID);
                break;
            //
            //phase II below
            //
            case "ACM":
                s_ACM ACM = new s_ACM();
                realImportNum = ACM.Execute(IiSheet, strGUID);
                break;
            case "ACS":
                s_ACS ACS = new s_ACS();
                realImportNum = ACS.Execute(IiSheet, strGUID);
                break;
            case "IEEE":
                s_IEEE IEEE = new s_IEEE();
                realImportNum = IEEE.Execute(IiSheet, strGUID);
                break;
            case "SDOL":
                s_SDOL SDOL = new s_SDOL();
                realImportNum = SDOL.Execute(IiSheet, strGUID);
                break;
            case "SPIE":
                s_SPIE SPIE = new s_SPIE();
                realImportNum = SPIE.Execute(IiSheet, strGUID);
                break;
            case "TI":
                s_TI TI = new s_TI();
                realImportNum = TI.Execute(IiSheet, strGUID);
                break;
            case "Wiley":
                s_Wiley Wiley = new s_Wiley();
                realImportNum = Wiley.Execute(IiSheet, strGUID);
                break;
            case "天下":
                s_TangShang TangShang = new s_TangShang();
                realImportNum = TangShang.Execute(IiSheet, strGUID);
                break;
            case "萬方":
                s_WangFang WangFang = new s_WangFang();
                realImportNum = WangFang.Execute(IiSheet, strGUID);
                break;
            case "通用":
                dt = _excel.ImportExecl(IiSheet);
                s_Utility Utility = new s_Utility();
                realImportNum = Utility.Execute(dt, strGUID, Year, Month);
                break;
            case "工程":
                dt = _excel.ImportExecl(IiSheet);
                s_Engineering Engineering = new s_Engineering();
                realImportNum = Engineering.Execute(dt, strGUID);
                break;
            case "聯合知識庫":
                dt = _excel.ImportExecl(IiSheet);
                s_Union Union = new s_Union();
                realImportNum = Union.Execute(dt, strGUID);
                break;
            case "SciFinder":
                dt = _excel.ImportExecl(IiSheet);
                s_SciFinder SciFinder = new s_SciFinder();
                realImportNum = SciFinder.Execute(dt, strGUID);
                break;
        }

        //每一次匯入都要做log,判斷等不等於零是因為沒資料就不要做log
        if (TextType.Trim() != "通用" && realImportNum != 0)
        {
            string yearMonth = Year + Month + "01";
            DateTime dTime = DateTime.ParseExact(yearMonth, "yyyyMMdd", CultureInfo.InvariantCulture);
            //通用的log自己在裡面做
            saveLog(realImportNum, strGUID, TextType, dTime);
        }

        //做完匯入之後再把工研人塞入
        UpdateEmpno(strGUID);

        //由於在通用的表單裡面,有些資料會有email,但是沒IP跟工號,
        //而也有可能是有工號沒IP沒EMAIL
        //所以要再等根據IP更新之後
        //2016/02/02 由於Goldfile有可能會用中文姓名+部門+單位來當WHERE條件,所以在這裡多一步去補上
        if ((TextType.Trim() == "通用" 
            || TextType.Trim() == "Goldfire" 
            || TextType.Trim() == "TI" 
            || TextType.Trim() == "聯合知識庫" 
            || TextType.Trim() == "SciFinder") && realImportNum != 0)
        {
            UpdateUtility(strGUID);
        }
        return realImportNum;
    }

    public static void AddParameters(SqlCommand oCmd)
    {
        oCmd.Parameters.Add("@main_parentid", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_ip", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_dbtype", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_viewC", SqlDbType.Int);
        oCmd.Parameters.Add("@main_downloadC", SqlDbType.Int);
        oCmd.Parameters.Add("@main_logC", SqlDbType.Int);
        oCmd.Parameters.Add("@main_indexC", SqlDbType.Int);
        oCmd.Parameters.Add("@main_datetime", SqlDbType.DateTime);
        oCmd.Parameters.Add("@main_title", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_chapter", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_account", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_empno", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_cname", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_email", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_org", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_orgcname", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_deptcd", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_deptcdcname", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_createdate", SqlDbType.DateTime);
        oCmd.Parameters.Add("@main_isbn", SqlDbType.NVarChar);
        oCmd.Parameters.Add("@main_issn", SqlDbType.NVarChar);
    }

    public static void RemoveExtraParameters(SqlCommand oCmd, bool c)
    {
        if (c)//做一遍就好省效能
        {
            int parametersCount = oCmd.Parameters.Count;
            for (int i = parametersCount - 1; i > 0; i--)
            {
                if (oCmd.Parameters[i].Value == null)
                {
                    oCmd.Parameters.RemoveAt(i);
                }
            }
        }
    }

    public void saveLog(int realImportNum, string strGUID, string dbtype, DateTime statistics_datetime)
    {
        int rowcount = realImportNum;

        SqlConnection oConn = new SqlConnection(AppConfig.DSN);
        oConn.Open();
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = oConn;
        //SqlTransaction myTrans = oConn.BeginTransaction();
        //oCmd.Transaction = myTrans;
        oCmd.Parameters.AddWithValue("@statistics_parentid", strGUID);
        oCmd.Parameters.AddWithValue("@statistics_dbtype", dbtype);
        oCmd.Parameters.AddWithValue("@statistics_num", rowcount);
        oCmd.Parameters.AddWithValue("@statistics_datetime", statistics_datetime);
        oCmd.CommandType = CommandType.Text;
        oCmd.CommandText = "INSERT INTO rs_statistics (statistics_parentid,statistics_dbtype,statistics_num,statistics_createdate,statistics_datetime) VALUES (@statistics_parentid,@statistics_dbtype,@statistics_num,GETDATE(),@statistics_datetime)";
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
        oConn.Close();
    }

    public void UpdateEmpno(string strGUID)
    {
        StringBuilder sb = new StringBuilder();
        SqlConnection oConn = new SqlConnection(AppConfig.DSN);
        oConn.Open();
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = oConn;
        //SqlTransaction myTrans = oConn.BeginTransaction();
        //oCmd.Transaction = myTrans;
        oCmd.Parameters.AddWithValue("@main_parentid", strGUID);
        oCmd.CommandType = CommandType.Text;
        oCmd.CommandTimeout = 999;
        sb.Append(@"UPDATE [dbo].[rs_main] SET ");
        sb.Append(@"main_empno = (SELECT Top 1 com_empno FROM opendatasource('SQLOLEDB','Data Source=ITRIDPO,2800;User ID=pubusagelog;Password=C59dCB4Rmd5G').pcinfo.dbo.v_ITRIPC_Log_His_usagelog WHERE pc_ip = main_ip),");
        sb.Append(@"main_cname = (SELECT Top 1 com_cname FROM opendatasource('SQLOLEDB','Data Source=ITRIDPO,2800;User ID=pubusagelog;Password=C59dCB4Rmd5G').pcinfo.dbo.v_ITRIPC_Log_His_usagelog WHERE pc_ip = main_ip),");
        sb.Append(@"main_email = (SELECT Top 1 com_mailadd FROM opendatasource('SQLOLEDB','Data Source=ITRIDPO,2800;User ID=pubusagelog;Password=C59dCB4Rmd5G').pcinfo.dbo.v_ITRIPC_Log_His_usagelog WHERE pc_ip = main_ip),");
        sb.Append(@"main_org = (SELECT Top 1 com_orgcd FROM opendatasource('SQLOLEDB','Data Source=ITRIDPO,2800;User ID=pubusagelog;Password=C59dCB4Rmd5G').pcinfo.dbo.v_ITRIPC_Log_His_usagelog WHERE pc_ip = main_ip),");
        sb.Append(@"main_orgcname = (SELECT Top 1 org_abbr_chnm2 FROM opendatasource('SQLOLEDB','Data Source=ITRIDPO,2800;User ID=pubusagelog;Password=C59dCB4Rmd5G').pcinfo.dbo.v_ITRIPC_Log_His_usagelog WHERE pc_ip = main_ip),");
        sb.Append(@"main_deptcd = (SELECT Top 1 com_deptcd FROM opendatasource('SQLOLEDB','Data Source=ITRIDPO,2800;User ID=pubusagelog;Password=C59dCB4Rmd5G').pcinfo.dbo.v_ITRIPC_Log_His_usagelog WHERE pc_ip = main_ip),");
        sb.Append(@"main_deptcdcname = (SELECT Top 1 dep_deptname FROM opendatasource('SQLOLEDB','Data Source=ITRIDPO,2800;User ID=pubusagelog;Password=C59dCB4Rmd5G').pcinfo.dbo.v_ITRIPC_Log_His_usagelog WHERE pc_ip = main_ip), ");
        sb.Append(@"main_posCname = (SELECT Top 1 posCname FROM opendatasource('SQLOLEDB','Data Source=ITRIDPO,2800;User ID=pubusagelog;Password=C59dCB4Rmd5G').pcinfo.dbo.v_ITRIPC_Log_His_usagelog WHERE pc_ip = main_ip), ");
        sb.Append(@"main_jobCname = (SELECT Top 1 jobCname FROM opendatasource('SQLOLEDB','Data Source=ITRIDPO,2800;User ID=pubusagelog;Password=C59dCB4Rmd5G').pcinfo.dbo.v_ITRIPC_Log_His_usagelog WHERE pc_ip = main_ip),");
        sb.Append(@"main_depCname = (SELECT Top 1 depCname FROM opendatasource('SQLOLEDB','Data Source=ITRIDPO,2800;User ID=pubusagelog;Password=C59dCB4Rmd5G').pcinfo.dbo.v_ITRIPC_Log_His_usagelog WHERE pc_ip = main_ip),");
        sb.Append(@"main_YearDiff = (SELECT Top 1 YearDiff FROM opendatasource('SQLOLEDB','Data Source=ITRIDPO,2800;User ID=pubusagelog;Password=C59dCB4Rmd5G').pcinfo.dbo.v_ITRIPC_Log_His_usagelog WHERE pc_ip = main_ip) ");
        //sb.Append(@"");
        //sb.Append(@"");
        sb.Append(@"WHERE main_parentid = @main_parentid AND main_ip is not null AND main_ip <> ''");
        //sb.Append(@"");
        oCmd.CommandText = sb.ToString();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
        oConn.Close();
    }


    public void UpdateUtility(string strGUID)
    {
        StringBuilder sb = new StringBuilder();
        SqlConnection oConn = new SqlConnection(AppConfig.DSN);
        oConn.Open();
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = oConn;
        //SqlTransaction myTrans = oConn.BeginTransaction();
        //oCmd.Transaction = myTrans;
        oCmd.Parameters.AddWithValue("@main_parentid", strGUID);
        oCmd.CommandType = CommandType.StoredProcedure;
        oCmd.CommandTimeout = 999;
        sb.Append(@"UpdateUtility");
        oCmd.CommandText = sb.ToString();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
        oConn.Close();
    }
}