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
                var nameSource = worksheetSource.Cells["C3"];
                var cellRange1 = worksheetSource.Cells["D6:D35"];
                var cellRange2 = worksheetSource.Cells["E6:E35"];
                var cellRange3 = worksheetSource.Cells["F6:F35"];

                foreach (var worksheet in excelFileDestination.Workbook.Worksheets)
                {
                    Console.WriteLine(worksheet);
                }

                var worksheetDestination = excelFileDestination.Workbook.Worksheets[0];

                var nameDestination = worksheetDestination.Cells["C5:I5"];
                var destination1 = worksheetDestination.Cells["D8:D37"];
                var destination2 = worksheetDestination.Cells["G8:G37"];
                var destination3 = worksheetDestination.Cells["J8:J37"];

                if (worksheetDestination.Cells["C5:I5"].Merge) {
                    worksheetDestination.Cells["C5:I5"].Merge = false;
                    nameSource.Copy(nameDestination);
                    worksheetDestination.Cells["C5:I5"].Merge = true;
                }

                foreach (var cell in destination1)
                {
                    string index = cell.ToString().Substring(1);
                    // Unmerge
                    string range1 = $"D{index}:F{index}";
                    if (worksheetDestination.Cells[range1].Merge)
                    {
                        worksheetDestination.Cells[range1].Merge = false;
                    }

                    string range2 = $"G{index}:I{index}";
                    if (worksheetDestination.Cells[range2].Merge)
                    {
                        worksheetDestination.Cells[range2].Merge = false;
                    }

                    string range3 = $"J{index}:L{index}";
                    if (worksheetDestination.Cells[range3].Merge)
                    {
                        worksheetDestination.Cells[range3].Merge = false;
                    }
                }

                cellRange1.Copy(destination1);
                cellRange2.Copy(destination2);
                cellRange3.Copy(destination3);

                foreach (var cell in destination1)
                {
                    string index = cell.ToString().Substring(1);

                    string range1 = $"D{index}:F{index}";
                    string range2 = $"G{index}:I{index}";
                    string range3 = $"J{index}:L{index}";

                    // Re-merge
                    if (!worksheetDestination.Cells[range1].Merge)
                    {
                        worksheetDestination.Cells[range1].Merge = true;
                    }

                    if (!worksheetDestination.Cells[range2].Merge)
                    {
                        worksheetDestination.Cells[range2].Merge = true;
                    }

                    if (!worksheetDestination.Cells[range3].Merge)
                    {
                        worksheetDestination.Cells[range3].Merge = true;
                    }
                }
                excelFileDestination.Save();
            }
        }
    }
}
