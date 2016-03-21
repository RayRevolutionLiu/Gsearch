using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.IO;

/// <summary>
/// Excel 的摘要描述
/// </summary>
public class Excel
{
	public Excel()
	{
		//
		// TODO: 在這裡新增建構函式邏輯
		//
	}

    int _CellIndexStart = 0;
    ISheet u_sheet = null;
    /// <summary>
    /// 匯出報表名稱
    /// </summary>
    string SheetName { get; set; }
    string Execlformat { get; set; }
    IWorkbook workbook;
    HttpPostedFile _PostedFile;
    string _Upload_physical_path;
    /// <summary>
    /// 2007 為 true 其他則 false
    /// </summary>
    /// <returns></returns>
    private bool ExeclXlsorXlsx()
    {
        bool success = false;
        if (Execlformat == "2007")
        {
            success = true;
        }
        return success;
    }

    private string File_extension()
    {
        if (ExeclXlsorXlsx())
        {
            return "xlsx";
        }
        else { return "xls"; }
    }


    /// <summary>
    /// 匯出DataTable 成 Execl
    /// </summary>
    /// <param name="gv">DataTable</param>
    /// <param name="Header">Execl用陣列</param>
    /// <param name="ExeclName">匯出的Execl名稱</param>
    /// <param name="format">版本_輸入字串:2003 or 2007</param>
    /// <param name="templet">是否要用樣版Execl 要的話 true 不要 則 false, 假如 false 的話 下面參數則無效</param>
    /// <param name="templet_PhysicalPath">Execl 樣版的實體路徑</param>
    /// <param name="InsertCells">從哪裡開始插入值</param>
    /// <param name="templet_SheetIndex">取哪一個 Sheet 起始值是 0</param>
    /// <param name="Insertion_position">需要傳入的值  { { Row, Cell, value }}</param>
    public IWorkbook ExportToFile(DataTable gv, string[] Header, string ExeclName, string format, bool templet, string templet_PhysicalPath, int InsertCells, int templet_SheetIndex)
    {
        try
        {
            int InsertIndex = 1;
            int InsertIndexHead = 0;
            if (templet) { if (Header != null) { InsertIndex = InsertCells + 1; } else { InsertIndex = InsertCells; } } //在哪裡插入內容
            if (templet) InsertIndexHead = InsertCells; //在哪裡插入表頭
            Execlformat = format;  //2003 or 2007
            SheetName = ExeclName; //匯出的名字
            #region 建立 WorkBook 及試算表

            ISheet mySheet1;
            if (!ExeclXlsorXlsx())
            {


                if (!templet)
                {
                    workbook = new HSSFWorkbook();
                    mySheet1 = (HSSFSheet)workbook.CreateSheet(SheetName);
                }
                else
                {
                    using (FileStream file = new FileStream(templet_PhysicalPath, FileMode.Open, FileAccess.Read))
                    {
                        workbook = new HSSFWorkbook(file);
                    }
                    mySheet1 = (HSSFSheet)workbook.GetSheetAt(templet_SheetIndex);
                }
            }
            else
            {
                if (!templet)
                {
                    workbook = new XSSFWorkbook();
                    mySheet1 = (XSSFSheet)workbook.CreateSheet(SheetName);
                }
                else
                {
                    workbook = new XSSFWorkbook(templet_PhysicalPath);
                    mySheet1 = (XSSFSheet)workbook.GetSheetAt(templet_SheetIndex);

                }
            }

            if (templet)
            {
                //if (Insertion_position != null)
                //{
                //    int Insertion_position_Row = 0;
                //    int Insertion_position_Cell = 0;

                //    int Insertion_position_MaxCount = Insertion_position.GetLength(0);
                //    int Insertion_position_index = 0;
                //    foreach (string Insertion_positionValue in Insertion_position)
                //    {
                //        switch (Insertion_position_index)
                //        {
                //            default:

                //                break;

                //            case 0:
                //                Insertion_position_Row = int.Parse(Insertion_positionValue);
                //                Insertion_position_index++;
                //                break;

                //            case 1:
                //                Insertion_position_Cell = int.Parse(Insertion_positionValue);
                //                Insertion_position_index++;
                //                break;

                //            case 2:

                //                mySheet1.GetRow(Insertion_position_Row).CreateCell(Insertion_position_Cell).SetCellValue(Insertion_positionValue);
                //                Insertion_position_index = 0;
                //                break;
                //        }


                //    }
                //}
            }

            #endregion

            #region 建立 sheet 內容
            if (Header != null)
            {
                // 建立 sheet 內容
                IRow rowHeader;
                if (!ExeclXlsorXlsx())
                {
                    rowHeader = (HSSFRow)mySheet1.CreateRow(0);
                }
                else { rowHeader = (XSSFRow)mySheet1.CreateRow(0); }
                // 建立 Header

                for (int i = InsertIndexHead, iCount = Header.Length; i < iCount; i++)
                {
                    string strValue = Header[i];
                    // ICellStyle cs = workbook.CreateCellStyle(); 
                    // cs.WrapText = true;
                    ICell cell = rowHeader.CreateCell(i);
                    cell.SetCellValue(strValue.Replace("&nbsp;", "").Trim());
                    //   cell.CellStyle = cs;
                }
            }
            // 建立 DataRow
            for (int i = 0, iCount = gv.Rows.Count; i < iCount; i++)
            {
                IRow rowItem;
                if (!ExeclXlsorXlsx())
                {
                    rowItem = (HSSFRow)mySheet1.CreateRow(i + InsertCells);
                }
                else { rowItem = (XSSFRow)mySheet1.CreateRow(i + InsertCells); }
                for (int j = 0, jCount = gv.Columns.Count; j < jCount; j++)
                {
                    string DatarowsText = gv.Rows[i][j].ToString();
                    ICell cell = rowItem.CreateCell(j);
                    ICellStyle cs = workbook.CreateCellStyle();

                    if (DatarowsText.ToString().IndexOf("\\n") >= 0)
                    {
                        // cs.Alignment = HorizontalAlignment.Left;
                        cs.WrapText = true;

                        rowItem.HeightInPoints = 2 * mySheet1.DefaultRowHeight / 20;
                    }
                    else
                    {
                        // cs.Alignment = HorizontalAlignment.Center;
                        // cs.VerticalAlignment = VerticalAlignment.Center;
                    }
                    cell.CellStyle = cs;
                    cell.SetCellValue(DatarowsText.Replace("\\r\\n", System.Environment.NewLine));
                }
            }
            #endregion


            int count = mySheet1.LastRowNum;
            for (int i = 0; i < count; i++)
            {
                mySheet1.AutoSizeColumn(i);
            }


            return workbook;
            //HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + HttpContext.Current.Server.UrlEncode(SheetName) + "." + File_extension()));
            //HttpContext.Current.Response.BinaryWrite(ms.ToArray());




        }
        catch (Exception ex)
        {
            throw ex;
        }
    }


    /// <summary>
    /// 匯出Execl
    /// </summary>
    public void ExportExeclfile()
    {
        // #region 匯出
        //FileStream sw = File.Create(HttpContext.Current.Server.MapPath(SheetName + "." + File_extension()));
        //workbook.Write(sw);
        //sw.Close();

        //System.IO.FileInfo fileinfo = new System.IO.FileInfo(HttpContext.Current.Server.MapPath(SheetName + "." + File_extension()));
        //if (fileinfo.Exists)
        //{
        //    HttpContext.Current.Response.ClearHeaders();
        //    HttpContext.Current.Response.Clear();
        //    string attachment = "attachment; filename=" + HttpContext.Current.Server.UrlEncode(SheetName) + "." + File_extension();
        //    HttpContext.Current.Response.AddHeader("Accept-Language", "zh-tw");
        //    HttpContext.Current.Response.AddHeader("Content-Disposition", attachment);
        //    HttpContext.Current.Response.AddHeader("Content-Length", fileinfo.Length.ToString());
        //    if (ExeclXlsorXlsx())
        //    {
        //        HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //    }
        //    else
        //    {
        //        HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
        //    }
        //    HttpContext.Current.Response.WriteFile(fileinfo.FullName);
        //    HttpContext.Current.Response.Flush();
        //    //刪除temp中該檔案
        //    fileinfo.Delete();
        //    HttpContext.Current.Response.End();
        //}
        // #endregion
        //#region 善後
        //workbook = null;
        //#endregion
        #region 匯出
        MemoryStream ms = new MemoryStream();
        //FileStream sw = new FileStream(SheetName + "." + File_extension(), FileMode.Create);
        workbook.Write(ms);
        HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("Attachment; Filename=" + HttpUtility.UrlEncode(SheetName) + "." + File_extension()));
        HttpContext.Current.Response.BinaryWrite(ms.ToArray());
        //釋放資源
        workbook = null;
        ms.Close();
        ms.Dispose();
        #endregion
    }
    /// <summary>
    /// 匯入Execl
    /// </summary>
    public DataTable ImportExecl(ISheet u_sheet)
    {
        DataTable VirTable = null;
        string excel_filePath = string.Empty;


        //excel_filePath = SaveFileAndReturnPath();//先上傳EXCEL檔案給Server
        //不同於Microsoft Object Model，NPOI都是從索引0開始算起
        //this.workbook = _IWorkbook();
        //this.u_sheet = _Sheet(workbook);  //取得第0個Sheet

        VirTable = GetExeclTable(u_sheet);
        return VirTable;
        // ShowError(Error);
    }

    private ISheet _Sheet(IWorkbook workbook)
    {
        if (Execl_Extension())
        {
            return (HSSFSheet)workbook.GetSheetAt(0);
        }
        else { return (XSSFSheet)workbook.GetSheetAt(0); }
    }
    private IWorkbook _IWorkbook()
    {
        if (Execl_Extension())
        {
            return new HSSFWorkbook(_PostedFile.InputStream);
        }
        else { return new XSSFWorkbook(_PostedFile.InputStream); }
    }
    /// <summary>
    /// 回傳副檔名 xls true , xlxs false
    /// </summary>
    /// <returns></returns>
    private bool Execl_Extension()
    {
        bool xls = true;
        if (System.IO.Path.GetExtension(_PostedFile.FileName).ToLower().Trim() == ".xlsx")
            xls = false;

        return xls;
    }

    #region 儲存EXCEL檔案給Server
    private string SaveFileAndReturnPath()
    {
        string return_file_path = "";//上傳的Excel檔在Server上的位置
        if (_PostedFile.FileName != "")
        {
            return_file_path = System.IO.Path.Combine(this._Upload_physical_path, Guid.NewGuid().ToString() + Path.GetFileName(_PostedFile.FileName));
            _PostedFile.SaveAs(return_file_path);
        }
        return return_file_path;
    }
    #endregion


    private DataTable GetExeclTable(ISheet u_sheet)
    {
        DataTable VirTable = new DataTable();
        string Record = string.Empty;
        int MaxRowArry = 0;

        if (u_sheet.LastRowNum > 0)
        {
            int rowArryIndex = 0;
            MaxRowArry = GetCellCount(u_sheet); //取得最大 cell 值

            //   MaxRowArry = (Cellend > MaxRowArry) ? MaxRowArry : Cellend;

            //Create Table
            //Record 則是用來記住Table Name 等一下把值塞進去
            for (rowArryIndex = 0; rowArryIndex < MaxRowArry; rowArryIndex++)
            {
                string RowName = Guid.NewGuid().ToString("D");
                VirTable.Columns.Add(RowName, System.Type.GetType("System.String"));
                if (!string.IsNullOrEmpty(Record)) Record += ",";
                Record += RowName;
            }
        }


        string[] RecordArry = Record.Split(',');
        int i = 0;
        int CellIndex = 0;
        //int CellCount = 0;
        // int LastRowNum = (ROWendIndex == 0) ? u_sheet.LastRowNum : (ROWendIndex > u_sheet.LastRowNum) ? u_sheet.LastRowNum : ROWendIndex;
        int LastRowNum = u_sheet.LastRowNum;

        for (i = u_sheet.FirstRowNum; i <= LastRowNum; i++)/*一列一列地讀取資料*/
        {
            IRow row = _Row(u_sheet, i);//取得目前的資料列
            int TheHeaderCellCount = MaxRowArry;

            if (row != null)
            {        
                //CellCount = row.Cells.Count;
                DataRow NewRow = VirTable.NewRow();
                if (_CellIndexStart != TheHeaderCellCount)
                {
                    for (CellIndex = _CellIndexStart; CellIndex < TheHeaderCellCount; CellIndex++)
                    {
                        if (row.GetCell(CellIndex) != null)
                        {
                            if (NewRow[RecordArry[CellIndex]] != null)
                            {
                                string cell0 = row.GetCell(CellIndex).ToString();
                                NewRow[RecordArry[CellIndex]] = cell0;
                            }
                        }
                    }
                }
                else
                {
                    if (row.GetCell(_CellIndexStart) != null)
                    {
                        if (NewRow[RecordArry[0]] != null)
                        {
                            string cell0 = row.GetCell(_CellIndexStart).ToString();
                            NewRow[RecordArry[CellIndex]] = cell0;
                        }
                    }
                }
                VirTable.Rows.Add(NewRow);
            }
        }

        return VirTable;
    }
    private IRow _Row(ISheet u_sheet, int i)
    {
        return u_sheet.GetRow(i);
    }

    private int GetCellCount(ISheet u_sheet)
    {
        int max = 0;
        int i = 0;
        if (u_sheet.LastRowNum > 0)
        {
            for (i = 0; i < u_sheet.LastRowNum; i++)
            {

                if (u_sheet.GetRow(i) != null)
                {
                    IRow rowArry = _Row(u_sheet, i);
                    if (max < rowArry.Cells.Count)
                        max = rowArry.Cells.Count;
                }
            }
        }
        return max;
    }
}