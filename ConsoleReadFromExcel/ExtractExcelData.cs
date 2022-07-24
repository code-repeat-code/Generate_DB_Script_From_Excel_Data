using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace ConsoleReadFromExcel
{
    class ExtractExcelData
    {
        readonly string fileName;
        readonly string tabName;

        public ExtractExcelData(string fileName, string tabName)
        {
            this.fileName = fileName;
            this.tabName = tabName;
        }

        public void readDataFromExcel()
        {
            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + @";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1'";
            OleDbConnection oledbConn = new OleDbConnection(connString);
            try
            {
                // Open connection
                oledbConn.Open();
                // Create OleDbCommand object and select data from worksheet Deummy Sheet excel sheet
                string query = "Select * FROM " + "[" + tabName + "$]";
                OleDbCommand cmd = new OleDbCommand(query, oledbConn);
                // Create new OleDbDataAdapter
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                oleda.SelectCommand = cmd;
                // Create a DataSet which will hold the data extracted from the worksheet.
                DataSet ds = new DataSet();
                oleda.Fill(ds);



                //Wrtie down you script here

                /*
                if (tabName == "Alternative Mutual Fund" || tabName == "Interval Funds")
                {
                    string vchTicker = "";
                    string vchFundName = "";
                    int vchMorningstarCat = 0;
                    int? bIsInterval = null;
                    if (tabName == "Interval Funds")
                    {
                        bIsInterval = 1;
                    }
                    
                    string recs = String.Empty;
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        vchTicker = (string)dr[0];
                        vchFundName = (string)dr[6];
                        foreach (DataColumn dc in ds.Tables[0].Columns)
                        { 
                            recs = $"IF NOT EXISTS(SELECT 1 FROM[dbo].[AltInvMutualFunds] WHERE vchTicker = '{vchTicker}')   INSERT INTO[dbo].[AltInvMutualFunds]([vchFundName],[vchTicker],[vchMorningstarCat],[bIsInterval]) VALUES('{vchFundName}', '{vchTicker}', '{vchMorningstarCat}', '{bIsInterval}')   GO";
                        }
                        writeToFile(recs);
                    }

                }
                if (tabName == "Leveraged_Inverse Funds" || tabName == "Inverse Funds") 
                {
                    string Symbol = string.Empty;
                    string Name = string.Empty;
                    string Leverage = "";
                    string ETFType = null;
                    int IsEnabled = 1;
                    DateTime CreatedDateTime = DateTime.Now;
                    string CreatedByName = "ADVISOR360";
                    DateTime? DeletedDateTime = null;
                    string DeletedByName = "NULL";
                    string recs = String.Empty;
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        Symbol = (string)dr[0];
                        Name = (string)dr[5];
                        foreach (DataColumn dc in ds.Tables[0].Columns)
                        {
                            recs = $"IF NOT EXISTS(SELECT 1 FROM [ComplianceAdministration].[ETFsymbols] WHERE [Symbol] = '{Symbol}')   INSERT INTO[ComplianceAdministration].[ETFsymbols] ([Symbol],[Name],[Leverage],[ETFType],[IsEnabled],[CreatedDateTime],[CreatedByName],[DeletedDateTime],[DeletedByName])   VALUES('{Symbol}', '{Name}', '{Leverage}', '{ETFType}', '{IsEnabled}', '{CreatedDateTime}', '{CreatedByName}', '{DeletedDateTime}', '{DeletedByName}')  GO";
                        }
                        writeToFile(recs);
                    }

                }
                */

            }
            catch (Exception e)
            {
                Console.WriteLine("Error :" + e.Message);
            }
            finally
            {
                // Close connection
                oledbConn.Close();
            }
        }

        //Write all the scripts the specific text file
        public void writeToFile(string records)
        {
            string scripts = records + Environment.NewLine;
            string textfilename = $"{tabName}.txt";
            string textfile = @"C:\Users\akumar\Documents\WorkStuff\" + textfilename;
            using (StreamWriter wr = new StreamWriter(textfile, true))
            {
                wr.Write(scripts);
            }

        }
    }
}
