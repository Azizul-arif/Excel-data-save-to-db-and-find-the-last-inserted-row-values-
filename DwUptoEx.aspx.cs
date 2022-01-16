using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace DWStock
{
    public partial class DwUptoEx : System.Web.UI.Page
    {
        protected void Upload(object sender, EventArgs e)
        {
            //Upload and save the file
            string excelPath = Server.MapPath("~/Files/") + Path.GetFileName(FileUpload1.PostedFile.FileName);
            FileUpload1.SaveAs(excelPath);

            string conString = string.Empty;
            string extension = Path.GetExtension(FileUpload1.PostedFile.FileName);
            switch (extension)
            {
                case ".xls": //Excel 97-03
                    conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                    break;
                case ".xlsx": //Excel 07 or higher
                    conString = ConfigurationManager.ConnectionStrings["Excel07+ConString"].ConnectionString;
                    break;

            }
            conString = string.Format(conString, excelPath);
            using (OleDbConnection excel_con = new OleDbConnection(conString))
            {
                excel_con.Open();
                string sheet1 = excel_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                DataTable dtExcelData = new DataTable();

                //[OPTIONAL]: It is recommended as otherwise the data will be considered as String by default.
                dtExcelData.Columns.AddRange(new DataColumn[16] { new DataColumn("Id", typeof(int)),
                new DataColumn("ItemCode", typeof(string)),
                new DataColumn("ItemName", typeof(string)),
                new DataColumn("ItemModel", typeof(string)),
                new DataColumn("Color", typeof(string)),
                new DataColumn("Quantity", typeof(float)),
                new DataColumn("EngineNo", typeof(string)),
                new DataColumn("ChassisNo",  typeof(string)),
                new DataColumn("Disc", typeof(string)),
                new DataColumn("PostingDate", typeof(DateTime)),
                new DataColumn("SystemDate", typeof(DateTime)),
                new DataColumn("Type", typeof(string)),
                new DataColumn("BatchNo", typeof(string)),
                new DataColumn("DealerCode", typeof(string)),
                new DataColumn("Remark", typeof(string)),
                new DataColumn("Price", typeof(float)),});

                using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [" + sheet1 + "]", excel_con))
                {
                    oda.Fill(dtExcelData);
                }
                excel_con.Close();

                string consString = ConfigurationManager.ConnectionStrings["DbConnection"].ConnectionString;
               
                using (SqlConnection con = new SqlConnection(consString))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                    {
                        //Set the database table name
                        sqlBulkCopy.DestinationTableName = "dbo.Stocks";

                        //[OPTIONAL]: Map the Excel columns with that of the database table
                        sqlBulkCopy.ColumnMappings.Add("Id", "Id");
                        sqlBulkCopy.ColumnMappings.Add("ItemCode", "ItemCode");
                        sqlBulkCopy.ColumnMappings.Add("ItemName", "ItemName");
                        sqlBulkCopy.ColumnMappings.Add("ItemModel", "ItemModel");
                        sqlBulkCopy.ColumnMappings.Add("Color", "Color");
                        sqlBulkCopy.ColumnMappings.Add("Quantity", "Quantity");
                        sqlBulkCopy.ColumnMappings.Add("EngineNo", "EngineNo");
                        sqlBulkCopy.ColumnMappings.Add("ChassisNo", "ChassisNo");
                        //if(sqlBulkCopy.ColumnMappings.Contains())
                        //{
                        //    throw new Exception("Data already Added");
                        //}
                        //else
                        //{
                        //    sqlBulkCopy.ColumnMappings.Add("ChassisNo", "ChassisNo");
                        //}
                        sqlBulkCopy.ColumnMappings.Add("Disc", "Disc");
                        sqlBulkCopy.ColumnMappings.Add("PostingDate", "PostingDate");
                        sqlBulkCopy.ColumnMappings.Add("SystemDate", "SystemDate");
                        sqlBulkCopy.ColumnMappings.Add("Type", "Type");
                        sqlBulkCopy.ColumnMappings.Add("DealerCode", "DealerCode");
                        sqlBulkCopy.ColumnMappings.Add("BatchNo", "BatchNo");
                        sqlBulkCopy.ColumnMappings.Add("Remark", "Remark");
                        sqlBulkCopy.ColumnMappings.Add("Price", "Price");
                        con.Open();
                        sqlBulkCopy.WriteToServer(dtExcelData);
                        con.Close();
                        //count data 

                        
                        
                    }
                }
                
                //alert after upload to db
                ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "alertMessage", "alert('Stock Inserted Successfully')", true);
                int numberOfRowsInserted = dtExcelData.Rows.Count;// <-- Count Row numbers
                
                string message=string.Format("<script>alert({0};</script>", numberOfRowsInserted);
                //ScriptManager.RegisterStartupScript(this, this.GetType(), "scr", message, false);
               // string text = "Total Data Insert";
                ScriptManager.RegisterStartupScript(this, this.GetType(), "scr",  message, false);

                //using (SqlConnection con = new SqlConnection(consString))
                //{

                //}
            }
        }
    }
}