﻿using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using WordOffice = Microsoft.Office.Interop.Word;
using ExcelOffice = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Data;
using ExcelDataReader;
using OPM.OPMEnginee;

namespace OPM.ExcelHandler
{
    class OpmExcelHandler : IExcelHandler
    {
        private string _strFileName;
         

        public string FileName
        {
            set { _strFileName = value; }
            get { return _strFileName; }
        }
        public OpmExcelHandler()
        { }
        ~OpmExcelHandler()
        { }
        
        public static int fReadExcelFile(string fname)
        {
            ExcelOffice.Range xlRange = null;
            ExcelOffice.Workbook xlWorkbook = null;
            ExcelOffice.Application xlApp = null;
            ExcelOffice._Worksheet xlWorksheet = null;
            try 
            {
                Dictionary<string, string> ListTP = new Dictionary<string, string>();

                xlApp = new ExcelOffice.Application();
                xlWorkbook = xlApp.Workbooks.Open(fname);
                xlWorksheet = (ExcelOffice._Worksheet)xlWorkbook.Sheets[2];
                xlRange = xlWorksheet.UsedRange;

                string xName = xlWorksheet.Name.ToString();
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                /*Get conntet Danh Sach Tinh/Tp*/
                for(int i=2; i<= rowCount; i++)
                {
                    string index = Convert.ToString((xlRange.Cells[i, 1] as ExcelOffice.Range).Text);
                    string strName = Convert.ToString((xlRange.Cells[i, 3] as ExcelOffice.Range).Text);
                    if(string.Empty != index && string.Empty != strName)
                    {
                        ListTP.Add(index, strName);
                    }else
                    {
                        i = rowCount + 1;
                    }    
                    
                }    
                //cleanup  
                 GC.Collect();
                GC.WaitForPendingFinalizers();
                //rule of thumb for releasing com objects:  
                //  never use two dots, all COM objects must be referenced and released individually  
                //  ex: [somthing].[something].[something] is bad  

                //release com objects to fully kill excel process from running in the background  
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release  
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release  
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return 1;
            }
            catch(Exception)
            {
                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //rule of thumb for releasing com objects:  
                //  never use two dots, all COM objects must be referenced and released individually  
                //  ex: [somthing].[something].[something] is bad  

                //release com objects to fully kill excel process from running in the background  
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release  
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release  
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return 0;
            }
            
            
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        public static int fReadExcelFile3(string fname)
        {

            ExcelOffice.Application xlApp = new ExcelOffice.Application();
            ExcelOffice.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            int numSheets = xlWorkbook.Sheets.Count;
            //ExcelOffice._Worksheet xlWorksheet = (ExcelOffice._Worksheet)xlWorkbook.Sheets[1];
            List<string> items = new List<string>();
            foreach (ExcelOffice._Worksheet xlWorksheet in xlWorkbook.Sheets)
            {

                string xName = xlWorksheet.Name.ToString();

                items.Add(xName);

            }
            //ExcelOffice.Range xlRange = xlWorksheet.UsedRange;
            //string xName = xlWorksheet.Name.ToString();

            //int rowCount = xlRange.Rows.Count;
            //int colCount = xlRange.Columns.Count;

            //cleanup  
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //rule of thumb for releasing com objects:  
            //  never use two dots, all COM objects must be referenced and released individually  
            //  ex: [somthing].[something].[something] is bad  

            //release com objects to fully kill excel process from running in the background  
            //Marshal.ReleaseComObject(xlRange);
            //Marshal.ReleaseComObject(xlWorksheet);

            //close and release  
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release  
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            return 1;
        }

        public static int fReadPackageListFiles(string[] fnames, ref List<Packagelist> oPackagelist)
        {
            ExcelOffice.Application xlApp = null;
            List<ExcelOffice.Workbook> xlWorkbookList = null;
            ExcelOffice.Range xlRange = null;
            ExcelOffice.Workbook xlWorkbook = null;
            ExcelOffice._Worksheet xlWorksheet = null;
            List<string> ListSerial = new List<string>();

            try
            {
                xlApp = new ExcelOffice.Application();
                xlWorkbookList = new List<ExcelOffice.Workbook>();
                foreach (string strfilename in fnames)
                {
                    int i = 0;
                    /*Read infor from file name*/
                    //int iRet = fReadInforFromFileName(strfilename, oPackagelist[i]);
                    //ExcelOffice.Application xlApp1;
                    //ExcelOffice.Workbook xlWorkBook1;
                    //ExcelOffice.Worksheet xlWorkSheet1;
                    //object misValue = System.Reflection.Missing.Value;

                    //xlApp1 = new ExcelOffice.Application();
                    //xlWorkBook1 = xlApp.Workbooks.Open(strfilename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    //xlWorkSheet1 = (ExcelOffice.Worksheet)xlWorkBook1.Worksheets.Item[1];

                    //MessageBox.Show(xlWorkSheet1.get_Range("A1", "A1").Value2.ToString());

                    //xlWorkBook1.Close(true, misValue, misValue);
                    //xlApp1.Quit();

                    //releaseObject(xlWorkSheet1);
                    //releaseObject(xlWorkBook1);
                    //releaseObject(xlApp1);


                    xlWorkbook = xlApp.Workbooks.Open(strfilename);
                    //xlWorkbook = xlApp.Workbooks.Open(strfilename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    xlWorksheet = (ExcelOffice._Worksheet)xlWorkbook.Sheets[3];
                    xlRange = xlWorksheet.UsedRange;
                    string xName = xlWorksheet.Name.ToString();
                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    /*Get conntet Danh Sach Tinh/Tp*/
                    for (int j = 2; j <= rowCount; j++)
                    {
                        //string strCase = Convert.ToString((xlRange.Cells[i, 1] as ExcelOffice.Range).Text);
                        //string strStorage = Convert.ToString((xlRange.Cells[i, 3] as ExcelOffice.Range).Text);
                        string strSerial = Convert.ToString((xlRange.Cells[j, 3] as ExcelOffice.Range).Text);
                        //string strMAC = Convert.ToString((xlRange.Cells[i, 3] as ExcelOffice.Range).Text);
                        //string strSerial_gpon = Convert.ToString((xlRange.Cells[i, 3] as ExcelOffice.Range).Text);
                        if (string.Empty != strSerial)
                        {
                            Packagelist temp = new Packagelist();
                            temp.SetSerial(strSerial);
                            int ret = fReadInforFromFileName(strfilename, temp);
                            oPackagelist.Add(temp);

                        }
                        else
                        {
                            j = rowCount + 1;
                        }

                    }
                    i++;
                }


                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //rule of thumb for releasing com objects:  
                //  never use two dots, all COM objects must be referenced and released individually  
                //  ex: [somthing].[something].[something] is bad  

                //release com objects to fully kill excel process from running in the background  
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release  
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release  
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return 1;
            }
            catch(Exception)
            {
                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //rule of thumb for releasing com objects:  
                //  never use two dots, all COM objects must be referenced and released individually  
                //  ex: [somthing].[something].[something] is bad  

                //release com objects to fully kill excel process from running in the background  
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release  
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release  
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return 0;
            }

            
        
        }

        public static int fReadInforFromFileName(string strFilename, Packagelist oPackagelist)
        {
            
            try 
            {
                    string[] strInfo = strFilename.Split("-");
                    string[] strDP = strInfo[6].Split(" ");
                    string[] strProvince = strInfo[8].Split(" ");
                    oPackagelist.PO_number = strDP[1];
                    oPackagelist.Province = strProvince[0];
                    oPackagelist.Year = strInfo[1] + strInfo[2];
                return 1;
            }
            catch(Exception)
            {
                return 0;
            }
        }

        public int fWriteExcelfile(string strFilename)
        {
            return 0;
        }

       
        public static  int getExcelSheet(ref DataSet result, string file, ComboBox cbx_sheet)
        {
            try
            {
                if (file.EndsWith(".xlsx"))
                {
                    // Reading from a binary Excel file (format; *.xlsx)
                    FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);
                    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    result = excelReader.AsDataSet();
                    excelReader.Close();
                }

                if (file.EndsWith(".xls"))
                {
                    // Reading from a binary Excel file ('97-2003 format; *.xls)
                    FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);
                    IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                    result = excelReader.AsDataSet();
                    excelReader.Close();
                }

                List<string> items = new List<string>();
                for (int i = 0; i < result.Tables.Count; i++)
                    items.Add(result.Tables[i].TableName.ToString());
                cbx_sheet.DataSource = items;
                return 1;
            }
            catch (Exception)
            {
                return 0;
            }


        }
    }
}
