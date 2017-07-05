using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.IO;

namespace XlxtoSql
{
    public partial class Default : System.Web.UI.Page
    {

        OleDbConnection OLDBEcon;
        SqlConnection SQLcon;

        string StrCon, StrQuery, Strsqlconn;

        protected void Page_Load(object sender, EventArgs e)
        {

        }
        //To create an oledbconnection for the Excel file
        private void ExcelConn(string FilePath)
        {

            try
            {
                StrCon = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", FilePath);
                OLDBEcon = new OleDbConnection(StrCon); 
            }
            catch (Exception ex)
            {
                ex.ToString();
            }

        }
        //To create an SQLconnection
        private void sqlconnection()
        {
            try
            {
                Strsqlconn = ConfigurationManager.ConnectionStrings["SqlCom"].ConnectionString;
                SQLcon = new SqlConnection(Strsqlconn);
            }
            catch (Exception ex)
            {
                ex.ToString();
                throw;
            }
        }

        //Create a function to read and insert an Excel File into the database
        private void InsertExcelRecords(string FilePath)
        {
            try
            {
                ExcelConn(FilePath);

                StrQuery = string.Format("Select [Timestamps],[WL, m],[ZG, m],[TL, m],[VExt, v],[VBatt, v] FROM [{0}]", "Sheet1");
                OleDbCommand Ecom = new OleDbCommand(StrQuery, OLDBEcon);

                  OLDBEcon.Open();

                DataSet ds = new DataSet();
                OleDbDataAdapter oda = new OleDbDataAdapter(StrQuery, OLDBEcon);
                OLDBEcon.Close();
                oda.Fill(ds);
                DataTable Exceldt = ds.Tables[0];
                sqlconnection();
                //creating object of SqlBulkCopy    
                SqlBulkCopy objbulk = new SqlBulkCopy(SQLcon);
                //assigning Destination table name    
                objbulk.DestinationTableName = "Weather";
                //Mapping Table column    
                objbulk.ColumnMappings.Add("Timestamps", "Timestamps");
                objbulk.ColumnMappings.Add("WL, m", "WL, m");
                objbulk.ColumnMappings.Add("ZG, m", "ZG, m");
                objbulk.ColumnMappings.Add("TL, m", "TL, m");
                objbulk.ColumnMappings.Add("VExt, v", "VExt, v");
                objbulk.ColumnMappings.Add("VBatt, v", "VBatt, v");
                //inserting Datatable Records to DataBase    
                SQLcon.Open();
                objbulk.WriteToServer(Exceldt);
                SQLcon.Close();
            }
            catch (Exception ex)
            {
                ex.ToString();
                throw;
            }

        }
        protected void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                string CurrentFilePath = Path.GetFullPath(FileUpload1.PostedFile.FileName);
                InsertExcelRecords(CurrentFilePath);
            }
            catch (Exception ex )
            {
                ex.ToString();
                throw;
            }
        }
    }
}