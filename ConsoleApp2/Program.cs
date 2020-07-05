using OfficeOpenXml;
using System;
using System.IO;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var fileSource = new FileInfo(@"C:\Users\nhatv\Work\VisualBasic\New-Data.xlsx");
            var fileDestination = new FileInfo(@"C:\Users\nhatv\Work\VisualBasic\Reports.xlsm");

            using (var excelFileSource = new ExcelPackage(fileSource))
            using (var excelFileDestination = new ExcelPackage(fileDestination))
            {
                var worksheetSource = excelFileSource.Workbook.Worksheets[0];

                var cellRange = worksheetSource.Cells["A2:D9"];

                var worksheetDestination = excelFileDestination.Workbook.Worksheets[0];

                var destination = worksheetDestination.Cells["A2"];

                cellRange.Copy(destination);

                excelFileDestination.Save();
            }
        }
    }
}
