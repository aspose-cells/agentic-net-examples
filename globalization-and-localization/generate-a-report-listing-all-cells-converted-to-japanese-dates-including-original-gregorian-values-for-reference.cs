// Title: Generate a Japanese era date report from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to list cells formatted with Japanese era dates in a new worksheet | Aspose.Cells .NET detect Japanese date format (年 月 日) and export original Gregorian values | Convert Gregorian DateTime cells to Japanese era using JapaneseCalendar in Aspose.Cells | Create a report of Japanese calendar cells from an existing Excel file with C# | Aspose.Cells iterate workbook and generate Japanese date conversion report
// Tags: Aspose.Cells Japanese era conversion report | C# detect Japanese date format Excel | Aspose.Cells iterate workbook cells | JapaneseCalendar formatting Aspose.Cells | Add report worksheet Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace JapaneseDateReportGenerator
{
    // The C# program loads an Excel workbook, adds a worksheet named "JapaneseDateReport", scans all other worksheets for DateTime cells whose custom number format contains Japanese characters (年, 月, 日), converts each Gregorian date to its Japanese era representation using JapaneseCalendar, and records the source sheet name, cell address, original Gregorian value, and the formatted Japanese date in the report sheet before autosizing columns and saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (replace with actual paths)
            string inputPath = @"C:\Input\SourceWorkbook.xlsx";
            string outputPath = @"C:\Output\Workbook_With_JapaneseDateReport.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Add a new worksheet for the report
                int reportSheetIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
                reportSheet.Name = "JapaneseDateReport";

                // Write header row
                reportSheet.Cells[0, 0].Value = "Source Sheet";
                reportSheet.Cells[0, 1].Value = "Cell Address";
                reportSheet.Cells[0, 2].Value = "Original Gregorian Value";
                reportSheet.Cells[0, 3].Value = "Japanese Date (Era)";

                int reportRow = 1; // Start writing data from the second row

                // Prepare Japanese culture with JapaneseCalendar
                CultureInfo jaCulture = new CultureInfo("ja-JP");
                jaCulture.DateTimeFormat.Calendar = new JapaneseCalendar();

                // Iterate through all worksheets except the report sheet itself
                for (int wsIdx = 0; wsIdx < workbook.Worksheets.Count; wsIdx++)
                {
                    if (wsIdx == reportSheetIndex) continue; // Skip the report sheet

                    Worksheet sheet = workbook.Worksheets[wsIdx];
                    Cells cells = sheet.Cells;

                    // Determine the used range to limit iteration
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];

                            // Check if the cell contains a DateTime value
                            if (cell.Value is DateTime originalDate)
                            {
                                // Retrieve the cell's custom number format string
                                string numberFormat = cell.GetStyle().Custom;

                                // Heuristic: Japanese date formats often contain Japanese characters like 年, 月, 日
                                if (!string.IsNullOrEmpty(numberFormat) &&
                                    (numberFormat.Contains("年") || numberFormat.Contains("月") || numberFormat.Contains("日")))
                                {
                                    // Convert the Gregorian date to Japanese era representation
                                    string japaneseDate = originalDate.ToString("gggee年M月d日", jaCulture);

                                    // Populate the report row
                                    reportSheet.Cells[reportRow, 0].Value = sheet.Name;
                                    reportSheet.Cells[reportRow, 1].Value = cell.Name; // e.g., "B5"
                                    reportSheet.Cells[reportRow, 2].Value = originalDate; // Gregorian value
                                    reportSheet.Cells[reportRow, 3].Value = japaneseDate;

                                    reportRow++;
                                }
                            }
                        }
                    }
                }

                // Auto-fit columns for better readability
                reportSheet.AutoFitColumns();

                // Save the workbook with the added report
                workbook.Save(outputPath);
                Console.WriteLine($"Report generated successfully at: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log or display the exception details for troubleshooting
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
