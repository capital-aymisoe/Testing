using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Model;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Model;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Data.SqlClient;
using System.Configuration;

namespace Testing.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpPost]
        public ActionResult FileUploadSave(HttpPostedFileBase uploadFile, Testing.Models.TestingModel model)
        {
            try

            {
                if (uploadFile != null)
                {
                    System.Data.DataTable dt = new System.Data.DataTable();
                    //dt = ExcelToTable(uploadFile.FileName);
                    //InitializeWorkbook(@"D:\JSAT\" + uploadFile.FileName);
                    //xlsxToDT();
                    //SampleData();
                    dt=TestingData(uploadFile.FileName);
                    UpdateCompanyTag(dt);
                }

            }

            catch(Exception ex) { string error = ex.ToString(); }

            return RedirectToAction("Index");
        }

        public DataTable TestingData(string excelfile)
        {
            string filename = @"D:\JSAT\" + excelfile;
            FileStream excelStream = new FileStream(filename, FileMode.Open);
            var book = new XSSFWorkbook(excelStream);          
            excelStream.Close();

            var sheet = book.GetSheetAt(2);
            var headerRow = sheet.GetRow(2);
            var cellCount = headerRow.LastCellNum;
            var rowCount = sheet.LastRowNum;//LastRowNum = PhysicalNumberOfRows - 1

            //header
            DataSet sampleDataSet = new DataSet();
            sampleDataSet.Locale = CultureInfo.InvariantCulture;
            DataTable table = sampleDataSet.Tables.Add("SampleData");

            table.Columns.Add("YYYYMM", typeof(string));
            table.Columns.Add("FingerPrintID", typeof(string));
            table.Columns.Add("StaffType", typeof(string));
            table.Columns.Add("AttandenceDate", typeof(string));
            table.Columns.Add("TimeIn", typeof(string));
            table.Columns.Add("TimeOut", typeof(string));

            string attdate = string.Empty;
            string newattdate = string.Empty;
            string yyymm = string.Empty;
            for (int j = sheet.FirstRowNum + 2; j <= 2; j++)
            {
                var row = sheet.GetRow(j);
                 attdate = GetCellValue(row.GetCell(2));
                string[] lines = Regex.Split(attdate, "-");
                 newattdate = lines[0] + "-" + lines[1];
                 yyymm = lines[0] + lines[1];
            }
            for (int j = sheet.FirstRowNum + 4; j <= Convert.ToUInt32(rowCount); j=j+2)
            {
                DataRow sampleDataRow;
                var row = sheet.GetRow(j);
               
                for (int i = row.FirstCellNum + 1; i <= Convert.ToUInt32(cellCount); i++)
                {
                    sampleDataRow = table.NewRow();

                    sampleDataRow["AttandenceDate"] = newattdate + "-" + i;

                    sampleDataRow["YYYYMM"] = yyymm;

                    sampleDataRow["FingerPrintID"] = 1;

                    if (j % 2 == 0)
                    {
                        var row1 = sheet.GetRow(j);
                        sampleDataRow["StaffType"] = GetCellValue(row1.GetCell(2));

                        var row2 = sheet.GetRow(j+1);
                        if (GetCellValue(row2.GetCell(i - 1)).ToString() != "")
                        {
                            sampleDataRow["TimeIn"] = GetCellValue(row2.GetCell(i - 1)).Substring(0, 5);
                        }
                        else
                        {
                            sampleDataRow["TimeIn"] = GetCellValue(row2.GetCell(i - 1));
                        }

                        if (GetCellValue(row2.GetCell(i - 1)).ToString() != "")
                        {
                            sampleDataRow["TimeOut"] = GetCellValue(row2.GetCell(i - 1)).Substring(GetCellValue(row2.GetCell(i - 1)).Length-5, 5);
                        }
                        else
                        {
                            sampleDataRow["TimeIn"] = GetCellValue(row2.GetCell(i - 1));
                        }
                    }
                    else
                    {
                        var row1 = sheet.GetRow(j-1);
                        sampleDataRow["StaffType"] = GetCellValue(row1.GetCell(2));

                        var row2 = sheet.GetRow(j);
                        if (GetCellValue(row2.GetCell(i - 1)).ToString() != "")
                        {
                            sampleDataRow["TimeIn"] = GetCellValue(row2.GetCell(i - 1)).Substring(0, 5); ;
                        }
                        else
                        {
                            sampleDataRow["TimeIn"] = GetCellValue(row2.GetCell(i - 1));
                        }
                        if (GetCellValue(row2.GetCell(i - 1)).ToString() != "")
                        {
                            sampleDataRow["TimeOut"] = GetCellValue(row2.GetCell(i - 1));
                        }
                        else
                        {
                            sampleDataRow["TimeOut"] = GetCellValue(row2.GetCell(i - 1)).Substring(GetCellValue(row2.GetCell(i - 1)).Length - 5, 5); 
                        }
                    }
                                       
                    table.Rows.Add(sampleDataRow);
                }

            }

            return sampleDataSet.Tables[0];
        }

        public void UpdateCompanyTag(DataTable dttest)
        {
            DataTable dt = new DataTable();
            SqlParameter[] prms = new SqlParameter[1];

            dttest.TableName = "Testing";
            System.IO.StringWriter writer = new System.IO.StringWriter();
            dttest.WriteXml(writer, XmlWriteMode.WriteSchema, false);
            string result = writer.ToString();
            prms[0] = new SqlParameter("@xml", SqlDbType.Xml) { Value = result };
            InsertUpdateDeleteData("M_Attandence_Insert", prms);
        }

        public static string conStr = ConfigurationManager.ConnectionStrings["JSAT_HRConnection"].ConnectionString;
        public void InsertUpdateDeleteData(string sSQL, params SqlParameter[] para)
        {
            var newCon = new SqlConnection(conStr);
            SqlCommand cmd = new SqlCommand(sSQL, newCon);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddRange(para);
            cmd.Connection.Open();
            cmd.ExecuteNonQuery();
            cmd.Connection.Close();
        }
        //private static DataSet SampleData()
        //{
        //    DataSet sampleDataSet = new DataSet();
        //    sampleDataSet.Locale = CultureInfo.InvariantCulture;
        //    DataTable sampleDataTable = sampleDataSet.Tables.Add("SampleData");

        //    sampleDataTable.Columns.Add("YYYYMM", typeof(string));
        //    sampleDataTable.Columns.Add("FingerPrintID", typeof(string));
        //    sampleDataTable.Columns.Add("StaffType", typeof(string));
        //    sampleDataTable.Columns.Add("AttandenceDate", typeof(string));
        //    sampleDataTable.Columns.Add("TimeIn", typeof(string));
        //    sampleDataTable.Columns.Add("TimeOut", typeof(string));
        //    DataRow sampleDataRow;
        //    for (int i = 1; i <= 49; i++)
        //    {
        //        sampleDataRow = sampleDataTable.NewRow();
        //        sampleDataRow["YYYYMM"] = "Cell1: " + i.ToString(CultureInfo.CurrentCulture);
        //        sampleDataRow["FingerPrintID"] = "Cell2: " + i.ToString(CultureInfo.CurrentCulture);
        //        sampleDataRow["StaffType"] = "Cell3: " + i.ToString(CultureInfo.CurrentCulture);
        //        sampleDataRow["AttandenceDate"] = "Cell4: " + i.ToString(CultureInfo.CurrentCulture);
        //        sampleDataRow["TimeIn"] = "Cell5: " + i.ToString(CultureInfo.CurrentCulture);
        //        sampleDataRow["TimeOut"] = "Cell6: " + i.ToString(CultureInfo.CurrentCulture);
        //        sampleDataTable.Rows.Add(sampleDataRow);
        //    }

        //    return sampleDataSet;
        //}
        public void InitializeWorkbook(string path)
        {
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                 XSSFWorkbook hssfworkbook=new XSSFWorkbook();
                hssfworkbook = new XSSFWorkbook(file);
            }
        }

        public void xlsxToDT()
        {
            DataTable dt = new DataTable();
            XSSFWorkbook hssfworkbook=new XSSFWorkbook();
            ISheet sheet = hssfworkbook.GetSheetAt(1);
            IRow headerRow = sheet.GetRow(0);
            IEnumerator rows = sheet.GetRowEnumerator();

            int colCount = headerRow.LastCellNum;
            int rowCount = sheet.LastRowNum;

            for (int c = 0; c < colCount; c++)
            {

                dt.Columns.Add(headerRow.GetCell(c).ToString());
            }

            bool skipReadingHeaderRow = rows.MoveNext();
            while (rows.MoveNext())
            {
                IRow row = (XSSFRow)rows.Current;
                DataRow dr = dt.NewRow();

                for (int i = 0; i < colCount; i++)
                {
                    ICell cell = row.GetCell(i);

                    if (cell != null)
                    {
                        dr[i] = cell.ToString();
                    }
                }
                dt.Rows.Add(dr);
            }

            hssfworkbook = null;
            sheet = null;
        }
        public DataTable ExcelToTable(string excelfile)
        {
            string filename = @"D:\JSAT\"+excelfile;
            FileStream excelStream = new FileStream(filename, FileMode.Open);
            var table = new System.Data.DataTable();
            var book = new XSSFWorkbook(excelStream);
            excelStream.Close();

            var sheet = book.GetSheetAt(0);
            var headerRow = sheet.GetRow(2);
            var cellCount = headerRow.LastCellNum;
            var rowCount = sheet.LastRowNum;//LastRowNum = PhysicalNumberOfRows - 1

            //header
            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                var column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                table.Columns.Add(column);
            }

            //body
            for (var i = sheet.FirstRowNum + 1; i <= rowCount; i++)
            {
                var row = sheet.GetRow(i);
                var dataRow = table.NewRow();
                if (row != null)
                {
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                            dataRow[j] = GetCellValue(row.GetCell(j));
                    }
                }
                table.Rows.Add(dataRow);
            }

            return table;
        }

        private static string GetCellValue(ICell cell)
        {


            if (cell == null)
                return string.Empty;
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return string.Empty;
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.Numeric:
                case CellType.Unknown:
                default:
                    return cell.ToString();//This is a trick to get the correct value of the cell. NumericCellValue will return a numeric value no matter the cell value is a date or a number
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Formula:
                    try
                    {
                        var e = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        return cell.ToString();
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString();
                    }
            }
        }
    }
    }