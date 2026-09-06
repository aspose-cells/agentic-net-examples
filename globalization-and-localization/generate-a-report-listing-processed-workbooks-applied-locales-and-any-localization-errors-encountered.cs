// Title: Generate a localization report for multiple Excel workbooks showing applied CultureInfo and any errors using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that loops through a collection of .xlsx file paths, sets workbook.Settings.CultureInfo to a specified locale, forces formula recalculation, saves each workbook, and creates a new workbook containing a worksheet with columns: Workbook Path, Applied Locale, Localization Error. | Refactor the program to read the list of workbook file paths and the target locale from an external JSON configuration file instead of hard‑coding them. | Adapt the solution so that the summary report is written to a CSV file rather than an Excel workbook, preserving the same three columns.
// Common Searches: asp.net apply french CultureInfo to Excel workbooks using Aspose.Cells | c# batch set locale on multiple .xlsx files and log errors | how to create a processing summary worksheet with Aspose.Cells in a console app | generate localization error report for Excel files in .NET | export Aspose.Cells summary report to CSV instead of XLSX
// Tags: apply CultureInfo to Aspose.Cells workbook | batch locale conversion for .xlsx files | capture localization exceptions in Aspose.Cells | generate Excel summary report with Aspose.Cells | export Aspose.Cells worksheet to CSV

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLocalizationReport
{
    // Holds information about each processed workbook
    // The sample iterates over a predefined list of Excel file paths, loads each workbook with Aspose.Cells, assigns a target locale (e.g., fr-FR) to workbook.Settings.CultureInfo, recalculates formulas, saves the workbook, and records any exception message. After processing, it creates a new workbook, adds a worksheet named "Localization Report", writes a header row and one row per file containing the workbook path, applied locale, and any error, auto‑fits the columns, and saves the report to a specified location. The console outputs the path of the generated report.
    class WorkbookProcessInfo
    {
        public string WorkbookPath { get; set; }
        public string AppliedLocale { get; set; }
        public string LocalizationError { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // List of workbook file paths to process
            List<string> workbookFiles = new List<string>
            {
                @"C:\Workbooks\Sample1.xlsx",
                @"C:\Workbooks\Sample2.xlsx",
                // Add more workbook paths as needed
            };

            // Desired locale to apply (example: French - France)
            string targetLocale = "fr-FR";

            // Collection to store processing results
            List<WorkbookProcessInfo> results = new List<WorkbookProcessInfo>();

            foreach (string filePath in workbookFiles)
            {
                var info = new WorkbookProcessInfo
                {
                    WorkbookPath = filePath,
                    AppliedLocale = targetLocale,
                    LocalizationError = string.Empty
                };

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Apply the locale to the workbook settings
                    workbook.Settings.CultureInfo = new CultureInfo(targetLocale);

                    // Optionally, you can force a recalculation to ensure locale‑dependent formulas are updated
                    workbook.CalculateFormula();

                    // Save the workbook back (overwrites original; adjust as needed)
                    workbook.Save(filePath);
                }
                catch (Exception ex)
                {
                    // Capture any localization related errors
                    info.LocalizationError = ex.Message;
                }

                results.Add(info);
            }

            // Create a new workbook for the report
            Workbook reportWorkbook = new Workbook();
            Worksheet sheet = reportWorkbook.Worksheets[0];
            sheet.Name = "Localization Report";

            // Write header row
            sheet.Cells[0, 0].PutValue("Workbook Path");
            sheet.Cells[0, 1].PutValue("Applied Locale");
            sheet.Cells[0, 2].PutValue("Localization Error");

            // Populate report rows
            for (int i = 0; i < results.Count; i++)
            {
                var r = results[i];
                int row = i + 1; // start after header
                sheet.Cells[row, 0].PutValue(r.WorkbookPath);
                sheet.Cells[row, 1].PutValue(r.AppliedLocale);
                sheet.Cells[row, 2].PutValue(r.LocalizationError);
            }

            // Auto‑fit columns for better readability
            sheet.AutoFitColumns();

            // Save the report workbook
            string reportPath = @"C:\Workbooks\LocalizationReport.xlsx";
            reportWorkbook.Save(reportPath);

            Console.WriteLine($"Localization report generated at: {reportPath}");
        }
    }
}
