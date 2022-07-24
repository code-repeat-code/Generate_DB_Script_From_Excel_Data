using System;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;

namespace ConsoleReadFromExcel
{
    class Program
    {


        public static void Main(String[] args)
        {
            string filename = @"C:\Users\akumar\Downloads\AI_Interval_Leveraged_Inverse Fund list from Product Master_with Fnd Names.xlsx";
            //string filename = FilePath;
            
            List<string> tabNamesList = new List<string> {
                "Alternative Mutual Fund","Interval Funds","Leveraged_Inverse Funds","Inverse Funds"
            };
            foreach (var tabName in tabNamesList) {
                ExtractExcelData ex = new ExtractExcelData(filename, tabName);
                ex.readDataFromExcel();
            }
            

        }
    }
}

