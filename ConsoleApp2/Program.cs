using OfficeOpenXml;
using System;
using System.IO;
using System.Reflection;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string pathDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            var fileSource = new FileInfo($"{pathDirectory}/Data/FlextimeSheetForm_202006_Akiyama.xlsx");
            var fileDestination = new FileInfo($"{pathDirectory}/Data/FlextimeSheetForm20200601.xlsx");

            using (var excelFileSource = new ExcelPackage(fileSource))
            using (var excelFileDestination = new ExcelPackage(fileDestination))
            {
                var worksheetSource = excelFileSource.Workbook.Worksheets[0];
                var cellRange = worksheetSource.Cells["D2:D35"];

                foreach (var worksheet in excelFileDestination.Workbook.Worksheets)
                {
                    Console.WriteLine(worksheet);
                }

                var worksheetDestination = excelFileDestination.Workbook.Worksheets[1];

                var destination = worksheetDestination.Cells["U8:U37"];

                cellRange.Copy(destination);

                excelFileDestination.Save();
            }
        }
    }
}
