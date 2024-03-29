﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Davids_Project
{
   public class Excel_Integration
    {
        private List<string> fileNameList;
        private int count;
        Dictionary<string, string> myDictionary = new Dictionary<string, string>();

        public async Dictionary<string, string> getExcelData(string myFile)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            bool excelWasRunning = System.Diagnostics.Process.GetProcessesByName("excel").Length > 0;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(myFile);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            fileNameList = new List<string>();

            try
            {
                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        //new line
                        if (j == 1)
                            Console.Write("\r\n");

                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

                        else if (xlRange.Cells[i, j].Value2.ToString() == "FriendlyName")
                        {
                            break;
                        }

                        if (xlRange.Cells[i, j].Value2.ToString() != null)
                        {
                            fileNameList.Add(xlRange.Cells[i, j].Value2.ToString());
                        }

                    }
                }
            }
            catch (Exception)
            {
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                if (xlApp != null) xlWorkbook.Close();
                if (xlApp != null) xlApp.Quit();

                xlApp = null;
            }

            var g = fileNameList.GroupBy(i => i);

            foreach (var grp in g)
            {
                Console.WriteLine("{0} {1}", grp.Key, grp.Count());
                myDictionary.Add(grp.Key, grp.Count().ToString());
            }

            //myGrid.Items.Add(myDictionary);

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close(true);

            xlApp = null;
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            if (xlApp != null) { xlApp.Quit(); }

            return myDictionary;
        }
    }
}
