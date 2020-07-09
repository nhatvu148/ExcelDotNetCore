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
                var cellRange = worksheetSource.Cells["D6:D35"];

                foreach (var worksheet in excelFileDestination.Workbook.Worksheets)
                {
                    Console.WriteLine(worksheet);
                }

                var worksheetDestination = excelFileDestination.Workbook.Worksheets[0];

                var destination = worksheetDestination.Cells["D8:D37"];

                foreach (var cell in destination)
                {
                    string range = cell.ToString() + ":" + cell.ToString().Substring(0, 1) + cell.ToString().Substring(1);
                    Console.WriteLine(range);
                }

                // cellRange.Copy(destination);

                excelFileDestination.Save();
            }
        }
    }
}
