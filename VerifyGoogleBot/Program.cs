using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using OfficeOpenXml;

namespace VerifyGoogleBot
{
    class Program
    {
        static void Main(string[] args)
        {
            var file = new FileInfo(@"C:\Users\deante\Desktop\NopErrorLogResults.xlsx");
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                int rowCount = worksheet.Dimension.Rows;
                int ColCount = worksheet.Dimension.Columns;

                var columnNumbers = new ErrorLogColumnNumbers();
                columnNumbers.IpAddress = LookFor(worksheet, "IpAddress");
                columnNumbers.Source = LookFor(worksheet, "Source");

                for (int row = 2; row <= rowCount; row++)
                {

                    var ipAddress = GetIntQ(worksheet, row, columnNumbers.IpAddress);

                    try
                    {
                        IPHostEntry hostEntry = Dns.GetHostEntry(ipAddress);
                        string domain = hostEntry.HostName;
                        worksheet.Cells[row, columnNumbers.Source].Value = domain;


                    }
                    catch (Exception)
                    {
                        worksheet.Cells[row, columnNumbers.Source].Value = "Failed";
                    }
                    
                }
                package.Save();
            }
        }

        public static string GetIntQ(OfficeOpenXml.ExcelWorksheet sheet, int row, int col)
        {
            var str = sheet.Cells[row, col].Value?.ToString();
            return String.IsNullOrWhiteSpace(str) ? null : str.Trim();
            // return str.Trim();
        }


        public static int LookFor(OfficeOpenXml.ExcelWorksheet sheet, string searchFor)
        {
            searchFor = searchFor.ToLower();
            for (var i = 1; i <= (sheet.Dimension?.Columns ?? 0); i++)
            {
                var cell = sheet.Cells[1, i];
                if (cell != null && cell.Text.ToLower() == searchFor)
                {
                    return i;
                }
            }
            throw new Exception("Cannot find column heading: '" + searchFor + "'");
        }

        private class ErrorLogColumnNumbers
        {
            public int IpAddress;
            public int Source;

        }
    }
}
