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
            // args[0] = Akiyama, args[1] = 20200601 args[2] = endRow
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string[] monthJPs = { "_1月", "_2月", "_3月", "_4月", "_5月", "_6月", "_7月", "_8月", "_9月", "_10月", "_11月", "_12月" };

            string pathDirectory = $"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}/..";
            var name = Convert.ToString(args[0]);
            var date = Convert.ToString(args[1]);
            var endRow = Convert.ToString(args[2]);

            bool isMonthJP = Int32.TryParse(date.Substring(4, 2), out int monthJP);

            string yearMonth = "";
            switch (monthJP)
            {
                case 12:
                    yearMonth = $"{date.Substring(0, 4)}12";
                    break;
                case 1:
                case 2:
                    yearMonth = $"{Int32.Parse(date.Substring(0, 4)) - 1}12";
                    break;
                case 3:
                case 4:
                case 5:
                    yearMonth = $"{date.Substring(0, 4)}03";
                    break;
                case 6:
                case 7:
                case 8:
                    yearMonth = $"{date.Substring(0, 4)}06";
                    break;
                case 9:
                case 10:
                case 11:
                    yearMonth = $"{date.Substring(0, 4)}09";
                    break;
                default:
                    Console.WriteLine("Default case");
                    break;
            }

            var fileSource = new FileInfo($"{pathDirectory}/output/FlextimeSheetForm_{date.Substring(0, 6)}_{name}.xlsx");
            var fileDestination = new FileInfo($"{pathDirectory}/FlextimeSheetForm{yearMonth}01.xlsx");
            var fileDestinationOutput = new FileInfo($"{pathDirectory}/output/FlextimeSheetForm_{date.Substring(0, 6)}_{name}.xlsx");

            using (var excelFileSource = new ExcelPackage(fileSource))
            using (var excelFileDestination = new ExcelPackage(fileDestination))
            {
                var worksheetSource = excelFileSource.Workbook.Worksheets[0];
                var nameSource = worksheetSource.Cells["C3"];
                var cellRange1 = worksheetSource.Cells[$"D6:D{endRow}"];
                var cellRange2 = worksheetSource.Cells[$"E6:E{endRow}"];
                var cellRange3 = worksheetSource.Cells[$"F6:F{endRow}"];

                string _monthJP = "";

                foreach (var worksheet in excelFileDestination.Workbook.Worksheets)
                {
                    if (worksheet.ToString().StartsWith(monthJPs[monthJP - 1]))
                    {
                        _monthJP = worksheet.ToString();
                    }
                }

                var worksheetDestination = excelFileDestination.Workbook.Worksheets[_monthJP];

                var nameDestination = worksheetDestination.Cells["C5:I5"];
                var destination1 = worksheetDestination.Cells[$"D8:D{Int32.Parse(endRow) + 2}"];
                var destination2 = worksheetDestination.Cells[$"G8:G{Int32.Parse(endRow) + 2}"];
                var destination3 = worksheetDestination.Cells[$"J8:J{Int32.Parse(endRow) + 2}"];

                if (worksheetDestination.Cells["C5:I5"].Merge)
                {
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

                excelFileDestination.SaveAs(fileDestinationOutput);
                // excelFileDestination.Save();
                Console.Write($"{name} downloaded FlextimeSheetForm_{date.Substring(0, 6)}_{name}.xlsx");
            }
        }
    }
}
