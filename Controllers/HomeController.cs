using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DWStock.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            string filePath = string.Empty;
            DataTable dt = new DataTable();
            
            if (file != null)
            {
                string path = Server.MapPath("~/Files/");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                filePath = path + Path.GetFileName(file.FileName);
                string extension = Path.GetExtension(file.FileName);
                file.SaveAs(filePath);

                string conString = string.Empty;

                switch (extension)
                {
                    case ".xls": //Excel 97-03.
                        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                        break;
                    case ".xlsx": //Excel 07 and above.
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                        break;
                }

                //DataTable dt = new DataTable();
                conString = string.Format(conString, filePath);

                using (OleDbConnection connExcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {
                            cmdExcel.Connection = connExcel;

                            //Get the name of First Sheet.
                            connExcel.Open();
                            DataTable dtExcelSchema;
                            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                            connExcel.Close();

                            //Read Data from First Sheet.
                            connExcel.Open();
                            cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                            odaExcel.SelectCommand = cmdExcel;
                            odaExcel.Fill(dt);
                            connExcel.Close();
                            var lastcolumn = dtExcelSchema.Rows.Count;
                            
                            
                        }
                    }
                }

                conString = @"Server=DESKTOP-4QOVSOF\SQLEXPRESS;Database=06_DEC;Trusted_Connection=True;";
                using (SqlConnection con = new SqlConnection(conString))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                    {
                        //Set the database table name.
                        sqlBulkCopy.DestinationTableName = "dbo.Stocks";

                        // Map the Excel columns with that of the database table, this is optional but good if you do
                        // 
                        //sqlBulkCopy.ColumnMappings.Add("Id", "Id");
                        sqlBulkCopy.ColumnMappings.Add("ItemCode", "ItemCode");
                        sqlBulkCopy.ColumnMappings.Add("ItemName", "ItemName");
                        sqlBulkCopy.ColumnMappings.Add("ItemModel", "ItemModel");
                        sqlBulkCopy.ColumnMappings.Add("Color", "Color");
                        sqlBulkCopy.ColumnMappings.Add("Quantity", "Quantity");
                        sqlBulkCopy.ColumnMappings.Add("EngineNo", "EngineNo");
                        sqlBulkCopy.ColumnMappings.Add("ChassisNo", "ChassisNo");
                        sqlBulkCopy.ColumnMappings.Add("Disc", "Disc");
                        sqlBulkCopy.ColumnMappings.Add("PostingDate", "PostingDate");
                        sqlBulkCopy.ColumnMappings.Add("SystemDate", "SystemDate");
                        sqlBulkCopy.ColumnMappings.Add("Type", "Type");
                        sqlBulkCopy.ColumnMappings.Add("DealerCode", "DealerCode");
                        sqlBulkCopy.ColumnMappings.Add("BatchNo", "BatchNo");
                        sqlBulkCopy.ColumnMappings.Add("Remark", "Remark");
                        sqlBulkCopy.ColumnMappings.Add("Price", "Price");

                        con.Open();
                        sqlBulkCopy.WriteToServer(dt);
                        con.Close();
                        
                    }
                }

                //int rowcount=dt.Rows.Count;
                //ViewBag.Message = rowcount;
                
            }
            //if the code reach here means everthing goes fine and excel data is imported into database
            int rowcount = dt.Rows.Count;//Count total no of  rows 
            ViewBag.Message = "Total data Inserted :" +rowcount + " Rows";
            ViewBag.Success = "File Imported and excel data saved into database";

            //find the last inserted column data
            var totaldatacount = dt.Columns.Count;
            var lastchassisno = dt.TableName.Contains("ChassisNo").ToString().LastOrDefault();
           // int lastColumn = sheet1.UsedRange.LastColumn;

            ViewBag.TotalDataCount = "Last Chassis Number: " + totaldatacount;
            

            //ViewBag.RowValues = "Row values" + rowval;

            var lastchassisnumber = dt.Rows[rowcount - 1][7].ToString(); //[7] is the index number of chassis column
            ViewBag.LastChassisNumber = "Value of last chassis: " + lastchassisnumber;


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
        public RedirectResult ExcelTest()
        {
            return Redirect("/DwUptoEx.aspx");
        }
    }
}