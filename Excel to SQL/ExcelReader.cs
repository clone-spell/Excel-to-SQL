using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Excel_to_SQL
{
    public class ExcelReader
    {
        public List<string> GetSheetNames(string filePath)
        {
            List<string> result = new List<string>();

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    foreach (ExcelWorksheet sheeet in package.Workbook.Worksheets)
                    {
                        result.Add(sheeet.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Get Sheet Names\n\nDescription :{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        public List<string> GetColumnNames(string filePath, string worksheetName)
        {
            List<string> result = new List<string>();

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorkbook excelWorkbook = package.Workbook;
                    ExcelWorksheet excelWorksheet = excelWorkbook.Worksheets[worksheetName];
                    for (int i = 1; i <= excelWorksheet.Dimension.End.Column; i++)
                    {
                        result.Add(excelWorksheet.Cells[1, i].Text);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Get Column Name\n\nDescription :{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return result;
        }

        public string GenerateInsertQueries(string excelFilePath, int[] colIndexs, string worksheetName, string kioskId, string officeName)
        {
            //set office code from kiosk id
            string officeCode = "";
            if (kioskId.Length >= 2)
                officeCode = kioskId.Substring(0, kioskId.Length - 2);

            StringBuilder queryBuilder = new StringBuilder();


            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetName]; 
                //error handling
                if (worksheet == null)
                {
                    MessageBox.Show("Worksheet is null", "Error", MessageBoxButtons.OK,MessageBoxIcon.Error);
                    return "";
                }
                try
                {
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        List<string> data = new List<string>();
                        foreach (int item in colIndexs)
                        {
                            string s = worksheet.Cells[row, item+1].Text;
                            data.Add(s);
                        }
                        string insertQuery = "";
                        if (data[3] != "" && data[2] != "")
                        {
                            insertQuery = "INSERT INTO TRANSACTIONS(CON_ID,CON_NO,NAME,ADDRESS1,ADDRESS2,OFF_NAME,OFF_CODE,PAYMENT_MODE,PAY_DT,CHE_NO,CHE_DT,MICR,BANK_CODE,BILLAMOUNT,AMOUNT,BATCHPAYMODEAMT,RCPTNO,BATCHNO,Denom1,Denom2,Denom3,Denom4,Denom5,Denom6,Denom7,Denom8,Denom9,Remarks,Trans_Process,BILL_Months,CHD,PaidBillMonth,ALREADYPAIDAMT,SERVERUPDATEDYN,KIOSKID,BillNo,SAPUPDATEDYN) values " +
                            $"({data[1]},'','','','','{officeName}','{officeCode}' ,'E','{FormatDate(data[0])}',0 ,'','','',{data[3]},{data[3]},{data[3]},{data[2]},{data[2]},0,0,0,0,0,0,0,0,0,'','ONLINE','{BillMonth(data[0])}','','{BillMonth(data[0])}',0,'Y','{kioskId}','','Y')" + Environment.NewLine;
                        }
                        else if (data[4] !="" && data[2] != "")
                        {
                            insertQuery = "INSERT INTO TRANSACTIONS(CON_ID,CON_NO,NAME,ADDRESS1,ADDRESS2,OFF_NAME,OFF_CODE,PAYMENT_MODE,PAY_DT,CHE_NO,CHE_DT,MICR,BANK_CODE,BILLAMOUNT,AMOUNT,BATCHPAYMODEAMT,RCPTNO,BATCHNO,Denom1,Denom2,Denom3,Denom4,Denom5,Denom6,Denom7,Denom8,Denom9,Remarks,Trans_Process,BILL_Months,CHD,PaidBillMonth,ALREADYPAIDAMT,SERVERUPDATEDYN,KIOSKID,BillNo,SAPUPDATEDYN) values " +
                            $"({data[1]},'','','','','{officeName}','{officeCode}' ,'Q','{FormatDate(data[0])}',0 ,'','','',{data[4]},{data[4]},{data[4]},{data[2]},{data[2]},0,0,0,0,0,0,0,0,0,'','ONLINE','{BillMonth(data[0])}','','{BillMonth(data[0])}',0,'Y','{kioskId}','','Y')" + Environment.NewLine;
                        }
                        else
                        {
                            insertQuery = "";
                        }
                        queryBuilder.Append(insertQuery);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Description: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                // Generate insert queries for each row
            }


            return queryBuilder.ToString();
        }

        public string FormatDate(string inputDate)
        {
            // Parse the input string into a DateTime object
            if (DateTime.TryParseExact(inputDate, "dd.MM.yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime date))
            {
                // Format the DateTime object into the desired format
                string formattedDate = date.ToString("MM/dd/yyyy") + " 00:00:00";
                return formattedDate;
            }
            else
            {
                // Handle parsing error (invalid input date format)
                return "Invalid date format";
            }
        }

        public string BillMonth(string input)
        {
            // Parse the input string into a DateTime object
            if (DateTime.TryParseExact(input, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
            {
                // Format the DateTime object into the desired format
                string formattedString = date.ToString("MMMyyyy").ToUpper();
                return formattedString;
            }
            else
            {
                // Handle parsing error (invalid input date format)
                return "Invalid date format";
            }
        }
        public ExcelReader() { }
    }
}
