// Title: Create a localization error report in Excel that lists workbook paths, applied CultureInfo locales, and processing exceptions using Aspose.Cells for .NET (C#)
// Common Searches: how to build an Excel localization report for multiple workbooks with Aspose.Cells in C# | Aspose.Cells set CultureInfo on a workbook and capture loading errors | C# generate a summary worksheet with workbook file path, locale and exception message using Aspose.Cells | auto‑fit columns and save a localization report workbook as .xlsx with Aspose.Cells | log localization processing failures to an Excel sheet using Aspose.Cells .NET API
// Tags: Aspose.Cells set workbook CultureInfo | Aspose.Cells log localization errors to worksheet | Aspose.Cells auto-fit columns for report sheet | Aspose.Cells save workbook as Xlsx format | Aspose.Cells generate Excel localization summary

using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells;

namespace LocalizationReportGenerator
{
    // The C# program iterates a list of workbook file paths with target locale identifiers, loads each workbook via Aspose.Cells, assigns the specified CultureInfo to the workbook's Settings, and records the file path, applied locale, and any exception message into a new worksheet. It auto‑fits the columns and saves the report as an Xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            // List of workbooks to process with their target locale identifiers.
            // Example entries; replace with actual data as needed.
            var workbooksToProcess = new List<(string FilePath, string Locale)>
            {
                (@"C:\Data\Workbook1.xlsx", "en-US"),
                (@"C:\Data\Workbook2.xlsx", "fr-FR"),
                (@"C:\Data\Workbook3.xlsx", "de-DE")
            };

            // Create a new workbook that will hold the report.
            var reportWorkbook = new Workbook();
            var reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "Localization Report";

            // Write header row.
            reportSheet.Cells[0, 0].PutValue("Workbook Path");
            reportSheet.Cells[0, 1].PutValue("Applied Locale");
            reportSheet.Cells[0, 2].PutValue("Localization Error");

            int currentRow = 1; // Start writing data from the second row.

            foreach (var (filePath, locale) in workbooksToProcess)
            {
                try
                {
                    // Load the workbook.
                    var workbook = new Workbook(filePath);

                    // Apply the requested locale.
                    var cultureInfo = new CultureInfo(locale);
                    workbook.Settings.CultureInfo = cultureInfo;

                    // Record successful processing.
                    reportSheet.Cells[currentRow, 0].PutValue(filePath);
                    reportSheet.Cells[currentRow, 1].PutValue(cultureInfo.Name);
                    reportSheet.Cells[currentRow, 2].PutValue(string.Empty);
                }
                catch (Exception ex)
                {
                    // Record any error that occurred during loading or locale application.
                    reportSheet.Cells[currentRow, 0].PutValue(filePath);
                    reportSheet.Cells[currentRow, 1].PutValue(locale);
                    reportSheet.Cells[currentRow, 2].PutValue(ex.Message);
                }

                currentRow++;
            }

            // Auto-fit columns for better readability.
            reportSheet.AutoFitColumns();

            // Save the report workbook.
            string reportPath = @"C:\Data\LocalizationReport.xlsx";
            reportWorkbook.Save(reportPath, SaveFormat.Xlsx);

            Console.WriteLine($"Localization report generated at: {reportPath}");
        }
    }
}
