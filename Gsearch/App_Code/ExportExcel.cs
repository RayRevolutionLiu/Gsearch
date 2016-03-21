using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using OfficeOpenXml;

/// <summary>
/// ExportExcel 的摘要描述
/// </summary>
public class ExportExcel
{
    private DataTable _dt;
    private string[] _header;

	public ExportExcel()
	{
        _dt = new DataTable();
	}

    public DataTable dt
    {
        get { return _dt; }
        set { _dt = value; }
    }
    public string[] header
    {
        get { return _header; }
        set { _header = value; }
    }


    public Stream CreateExcel(List<Part> part)
    {
        try
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            //HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet mySheet1;
            string SheetName = DateTime.Now.ToString("yyyyMMdd");
            mySheet1 = (XSSFSheet)workbook.CreateSheet(SheetName);
            int RowHeader = 0;

            if (part != null)
            {
                foreach (Part aPart in part)
                {
                    IRow PeriousRow = (XSSFRow)mySheet1.CreateRow(part.IndexOf(aPart));
                    ICell cell = PeriousRow.CreateCell(1);
                    cell.SetCellValue(aPart.ToString());
                    //Console.WriteLine(aPart);
                }

                RowHeader += part.Count;
            }

            
            IRow rowHeader = (XSSFRow)mySheet1.CreateRow(RowHeader);
            //塞入header的值
            for (int i = 0; i < header.Length; i++)
            {
                string strValue = header[i];
                // ICellStyle cs = workbook.CreateCellStyle(); 
                // cs.WrapText = true;
                ICell cell = rowHeader.CreateCell(i);
                cell.SetCellValue(strValue.Replace("&nbsp;", "").Trim());
                //   cell.CellStyle = cs;
            }

            //塞入底下的內容
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow rowItem;
                rowItem = (XSSFRow)mySheet1.CreateRow(i + RowHeader + 1);
                //rowItem = (HSSFRow)mySheet1.CreateRow(i + InsertCells);

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    string DatarowsText = dt.Rows[i][j].ToString();
                    ICell cell = rowItem.CreateCell(j);
                    //ICellStyle cs = workbook.CreateCellStyle();

                   // if (DatarowsText.ToString().IndexOf("\\n") >= 0)
                    //{
                        // cs.Alignment = HorizontalAlignment.Left;
                    //    cs.WrapText = true;

                    //    rowItem.HeightInPoints = 2 * mySheet1.DefaultRowHeight / 20;
                    //}
                    //else
                    //{
                        // cs.Alignment = HorizontalAlignment.Center;
                        // cs.VerticalAlignment = VerticalAlignment.Center;
                    //}
                   // cell.CellStyle = cs;
                    cell.SetCellValue(DatarowsText.Replace("\\r\\n", System.Environment.NewLine));
                }
            }


            //int count = dt.Columns.Count;
            //for (int i = 0; i < count; i++)
            //{
            //    mySheet1.AutoSizeColumn(i);
            //}

            #region 匯出
            
            MemoryStream ms = new MemoryStream();
            //FileStream sw = new FileStream(SheetName + "." + File_extension(), FileMode.Create);
            workbook.Write(ms);
            HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("Attachment; Filename=" + HttpUtility.UrlEncode(SheetName) + ".xlsx"));
            HttpContext.Current.Response.BinaryWrite(ms.ToArray());
            //釋放資源
            workbook = null;
            //ms.Close();
            //ms.Dispose();
            #endregion

            return ms;
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }


    /*底下是用新的工具匯出Excel,不過一樣會跳出錯誤,不過Excel的內容卻正確 */
    //public Stream CreatExcelOpenXml()
    //{
    //    //在記憶體中建立一個Excel物件
    //    ExcelPackage ep = new ExcelPackage();
    //    string SheetName = DateTime.Now.ToString("yyyyMMdd");

    //    //加入一個Sheet
    //    ep.Workbook.Worksheets.Add(SheetName);
    //    //取得剛剛加入的Sheet(實體Sheet就叫MySheet)

    //    ExcelWorksheet sheet1 = ep.Workbook.Worksheets[SheetName];//取得Sheet1 

    //    //塞入header的值
    //    for (int i = 0; i < header.Length; i++)
    //    {
    //        string strValue = header[i];
    //        sheet1.Cells[1, i + 1].Value = strValue;//加入標頭
    //    }

    //    #region 匯出

    //    MemoryStream ms = new MemoryStream();
    //    //FileStream sw = new FileStream(SheetName + "." + File_extension(), FileMode.Create);
    //    ep.SaveAs(ms);
    //    //workbook.Write(ms);
    //    HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("Attachment; Filename=" + HttpUtility.UrlEncode(SheetName) + ".xlsx"));
    //    HttpContext.Current.Response.BinaryWrite(ms.ToArray());
    //    //釋放資源
    //    ep = null;
    //    //ms.Close();
    //    //ms.Dispose();
    //    #endregion

    //    return ms;
    //}
}


public class Part : IEquatable<Part>
    {
        public string title { get; set; }

        public string values { get; set; }

        public override string ToString()
        {
            return title+":"+values;
        }

        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            Part objAsPart = obj as Part;
            if (objAsPart == null) return false;
            else return Equals(objAsPart);
        }

        public bool Equals(Part other)
        {
            if (other == null) return false;
            return (this.title.Equals(other.title));
        }

    // Should also override == and != operators.

    }