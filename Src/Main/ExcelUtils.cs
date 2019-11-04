using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using DataTable = System.Data.DataTable;

namespace USC.GISResearchLab.Common.Utils.Excel
{
    /// <summary>
    /// Summary description for ExcelUtils.
    /// </summary>
    public class ExcelUtils
    {
        public ExcelUtils()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public static DataTable ReadExcel2000FirstSheet(string fileName)
        {
            DataTable ret = new DataTable();

            string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=Excel 8.0";
            OleDbConnection conn = new OleDbConnection(connString);
            try
            {
                conn.Open();

                // get the name of the first sheet
                DataTable dataTableSheets = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheet = dataTableSheets.Rows[0]["TABLE_NAME"].ToString();

                // now get the data off the first sheet
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter();
                dataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM [" + sheet + "]", conn);
                dataAdapter.Fill(ret);
            }
            catch (Exception e)
            {
                string error = e.Message;
                throw new Exception("Error occured reading from Excel Sheet '" + fileName + "': " + error, e);
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }

            return ret;
        }

        public static DataTable ReadExcel2007FirstSheet(string fileName)
        {
            return ReadExcel2007FirstSheet(fileName, false, true);
        }


        // if there are charts in the xslx file, this will crash unless a hotfix is installed: http://support.microsoft.com/default.aspx?scid=kb;EN-US;968861 
        public static DataTable ReadExcel2007FirstSheet(string fileName, bool firstRowHeader, bool readAllAsText)
        {
            DataTable ret = new DataTable();


            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0 Xml";

            if (firstRowHeader)
            {
                connString += ";HDR=NO";
            }
            else
            {
                connString += ";HDR=YES";
            }

            if (readAllAsText)
            {
                connString += ";IMEX=1";
            }

            connString += "\"";

            OleDbConnection conn = new OleDbConnection(connString);
            try
            {
                conn.Open();

                // get the name of the first sheet
                DataTable dataTableSheets = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                string sheet = "";
                foreach (DataRow row in dataTableSheets.Rows)
                {
                    string currentSheet = row["TABLE_NAME"].ToString();

                    if (!currentSheet.ToLower().Contains("chart"))
                    {
                        sheet = currentSheet;
                        break;
                    }
                }

                if (!String.IsNullOrEmpty(sheet))
                {
                    // now get the data off the first sheet
                    OleDbDataAdapter dataAdapter = new OleDbDataAdapter();
                    dataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM [" + sheet + "]", conn);
                    dataAdapter.Fill(ret);
                }
            }
            catch (Exception e)
            {
                string error = e.Message;
                throw new Exception("Error occured reading from Excel Sheet '" + fileName + "': " + error, e);
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }

            return ret;
        }

        public static List<DataTable> ReadExcel2007Sheets(string fileName, bool firstRowHeader, bool readAllAsText)
        {
            List<DataTable> ret = new List<DataTable>();



            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0 Xml";

            if (firstRowHeader)
            {
                connString += ";HDR=NO";
            }
            else
            {
                connString += ";HDR=YES";
            }

            if (readAllAsText)
            {
                connString += ";IMEX=1";
            }

            connString += "\"";

            OleDbConnection conn = new OleDbConnection(connString);
            try
            {
                conn.Open();

                // get the name of the first sheet
                DataTable dataTableSheets = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                string sheet = "";
                foreach (DataRow row in dataTableSheets.Rows)
                {
                    string currentSheet = row["TABLE_NAME"].ToString();

                    if (!currentSheet.ToLower().Contains("chart"))
                    {
                        sheet = currentSheet;

                        if (!String.IsNullOrEmpty(sheet))
                        {
                            // now get the data off the first sheet
                            OleDbDataAdapter dataAdapter = new OleDbDataAdapter();
                            dataAdapter.SelectCommand = new OleDbCommand("SELECT * FROM [" + sheet + "]", conn);
                            DataTable dataTable = new DataTable();
                            dataAdapter.Fill(dataTable);

                            ret.Add(dataTable);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                string error = e.Message;
                throw new Exception("Error occured reading from Excel Sheet '" + fileName + "': " + error, e);
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }

            return ret;
        }

        public static void ConvertHTMLTableToExcel2000(string fileName)
        {
            Application xlApp = new Application();
            //xlApp.Workbooks.Open(@"C:\Projects\mamour\AggValueHistory\default.htm", null);
            Workbook theWorkbook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlApp.ActiveWorkbook.SaveAs(fileName + ".xls", -4143, "", "", false, false, XlSaveAsAccessMode.xlShared, false, true, false, false, false);
            if (xlApp != null)
            {
                xlApp.ActiveWorkbook.Close(false, fileName, false);
                xlApp.Quit();
                xlApp = null;
            }
        }
    }
}
