using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office;
using System.Runtime.InteropServices;//added Microsoft.Office.Interop.Excel assembly
using System.Data.OleDb;
using System.Data;
using System.Diagnostics;
using System.ComponentModel;
using System.IO;


namespace CRM.Models
{
    public class ExcelInfo
    {
        public ExcelInfo(string strFileName)
        {
            FileName = strFileName;
        }
        public DataSet ExcelDataSet { get; set; }
        public Exception LastException { get; set; }
        private string[] Extensions = new string[] { ".xls", ".xlsx" };
        private string mstrFileName;
        public string FileName
        {
            get { return mstrFileName; }
            set
            {
                if (!Extensions.Contains(System.IO.Path.GetExtension(value.ToLower())))
                    throw new Exception("Invalid file name");
                if (!File.Exists(value))
                    throw new Exception("File Not Found");
                mstrFileName = value;
            }
        }
        private List<string> mNameRanges = new List<string>();
        public List<string> NameRanges { get { return mNameRanges; } }

        private List<string> mSheets = new List<string>();
        public List<string> Sheets { get { return mSheets; } }


        public bool GetInformation()
        {
            Application xlApp = null;
            Workbooks xlWorkBooks = null;
            Workbook xlWorkBook = null;
            Workbook xlActiveRanges = null;
            Names xlNames = null;
            Sheets xlWorkSheets = null;
            bool Success = true;



            if (!System.IO.File.Exists(FileName))
            {
                Exception objEx = new Exception("Failed to locate '" + FileName + "'");
                LastException = objEx;
                throw objEx;
            }

            mSheets.Clear();
            mNameRanges.Clear();

            try
            {
                xlApp = new Application();
                xlApp.DisplayAlerts = false;
                xlWorkBooks = xlApp.Workbooks;
                xlWorkBook = xlWorkBooks.Open(FileName);

                xlActiveRanges = xlApp.ActiveWorkbook;
                xlNames = xlActiveRanges.Names;

                for (int x = 1; x <= xlNames.Count; x++)
                {
                    Name xlName = xlNames.Item(x);
                    mNameRanges.Add(xlName.Name);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlName);
                    xlName = null;
                }

                xlWorkSheets = xlWorkBook.Sheets;

                for (int x = 1; x <= xlWorkSheets.Count; x++)
                {
                    Worksheet Sheet1 = (Worksheet)xlWorkSheets[x];
                    mSheets.Add(Sheet1.Name);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Sheet1);
                    Sheet1 = null;
                }

                xlWorkBook.Close();
                xlApp.UserControl = true;
                xlApp.Quit();
                TryKillProcessByMainWindowHwnd(xlApp.Hwnd);//I kept getting excel hung up.This method forces the app to close...
            }
            catch (Exception objEx)
            {
                LastException = objEx;
                Success = false;
            }
            finally
            {
                if (xlWorkSheets != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheets);
                    xlWorkSheets = null;
                }

                if (xlNames != null)
                {
                    Marshal.FinalReleaseComObject(xlNames);
                    xlNames = null;
                }

                if (xlActiveRanges != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlActiveRanges);
                    xlActiveRanges = null;
                }

                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }

                if (xlWorkBooks != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBooks);
                    xlWorkBooks = null;
                }

                if (xlApp != null)
                {
                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return Success;
        }

        public void LoadDataFromExcel(string strRangeName)
        {
            String sConnectionString;
            OleDbConnection objConn;
            OleDbCommand objCmdSelect;
            OleDbDataAdapter objAdapter;


            sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                        "Data Source=" + FileName + ";" +
                        "Extended Properties='Excel 12.0;IMEX=1;HDR=NO;ImportMixedTypes=Text;TypeGuessRows=0;'";

            objConn = new OleDbConnection(sConnectionString);
            objConn.Open();

            // Create new OleDbCommand to return data from worksheet.
            objCmdSelect = new OleDbCommand("SELECT * FROM " + strRangeName, objConn);

            // Create new OleDbDataAdapter that is used to build a DataSet
            // based on the preceding SQL SELECT statement.
            objAdapter = new OleDbDataAdapter();

            // Pass the Select command to the adapter.
            objAdapter.SelectCommand = objCmdSelect;

            // Create new DataSet to hold information from the worksheet.
            ExcelDataSet = new DataSet();

            // Fill the DataSet with the information from the worksheet.
            objAdapter.Fill(ExcelDataSet, "XLData");

            // Clean up objects.
            objConn.Close();

            BuildHeadersFromFirstRowThenRemoveFirstRow();
        }

        private void BuildHeadersFromFirstRowThenRemoveFirstRow()
        {
            System.Data.DataTable dt = ExcelDataSet.Tables[0];
            DataRow firstRow = dt.Rows[0];


            for (int i = 0; i < dt.Columns.Count; i++)
            {
                dt.Columns[i].ColumnName = firstRow[i].ToString().Trim();
            }

            dt.Rows.RemoveAt(0);
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        /// <summary> Tries to find and kill process by hWnd to the main window of the process.</summary> 
        /// <param name="hWnd">Handle to the main window of the process.</param> 
        /// <returns>True if process was found and killed. False if process was not found by hWnd or if it could not be killed.</returns> 
        public static bool TryKillProcessByMainWindowHwnd(int hWnd)
        {
            uint processID;

            GetWindowThreadProcessId((IntPtr)hWnd, out processID);
            if (processID == 0)
                return false;

            try
            {
                Process.GetProcessById((int)processID).Kill();
            }
            catch (ArgumentException)
            {
                return false;
            }
            catch (Win32Exception)
            {
                return false;
            }
            catch (NotSupportedException)
            {
                return false;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
            return true;
        }  /// <summary> Finds and kills process by hWnd to the main window of the process.</summary> 
        /// <param name="hWnd">Handle to the main window of the process.</param> /// <exception cref="ArgumentException"> 
        /// Thrown when process is not found by the hWnd parameter (the process is not running).  
        /// The identifier of the process might be expired. /// </exception> 
        /// <exception cref="Win32Exception">See Process.Kill() exceptions documentation.</exception> 
        /// <exception cref="NotSupportedException">See Process.Kill() exceptions documentation.</exception> 
        /// <exception cref="InvalidOperationException">See Process.Kill() exceptions documentation.</exception> 
        public static void KillProcessByMainWindowHwnd(int hWnd)
        {
            uint processID;
            GetWindowThreadProcessId((IntPtr)hWnd, out processID);
            if (processID == 0)
                throw new ArgumentException("Process has not been found by the given main window handle.", "hWnd");
            Process.GetProcessById((int)processID).Kill();
        }
    }
}
