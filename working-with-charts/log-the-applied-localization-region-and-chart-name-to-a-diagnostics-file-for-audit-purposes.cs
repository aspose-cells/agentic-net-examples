// Title: Append workbook localization region and chart identifiers to a diagnostics log using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, reads the workbook's CultureInfo, and appends the region name to a diagnostics file. | Develop a method that walks through every worksheet and chart, supplies a fallback name for unnamed charts, and writes worksheet name, chart index, and chart name to a log file.
// Common Searches: Aspose.Cells how to write workbook culture info to a log file in C# | C# log all chart names from an Excel workbook using Aspose.Cells | record Excel chart index and worksheet name for audit with Aspose.Cells .NET | append diagnostics log with localization region and chart details Aspose.Cells | handle empty chart names when exporting chart information using Aspose.Cells
// Tags: append workbook cultureinfo to log Aspose.Cells | iterate worksheets and charts Aspose.Cells | log chart index and name .NET | fallback chart name generation Aspose.Cells | create diagnostics file for Excel audit

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example demonstrates how to load an Excel workbook with Aspose.Cells, retrieve the workbook's applied CultureInfo, and write that region together with each chart's identifier (using a generated fallback when the name is empty) to a diagnostics log file, creating the log directory if necessary.
class ChartAuditLogger
{
    // Path to the Excel file to be processed
    private const string InputFilePath = @"C:\Data\Report.xlsx";

    // Path to the diagnostics log file
    private const string LogFilePath = @"C:\Logs\ChartAudit.log";

    static void Main()
    {
        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(InputFilePath))
            {
                Console.Error.WriteLine($"Input file not found: {InputFilePath}");
                return;
            }

            // Ensure the log directory exists
            string logDir = Path.GetDirectoryName(LogFilePath);
            if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }

            // Load the workbook inside a safe block
            Workbook workbook;
            try
            {
                workbook = new Workbook(InputFilePath);
            }
            catch (Exception loadEx)
            {
                Console.Error.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Determine the applied localization region (culture)
            CultureInfo culture = workbook.Settings.CultureInfo;
            string region = culture?.Name ?? "Invariant";

            // Prepare the log file (append if it exists)
            using (StreamWriter logWriter = new StreamWriter(LogFilePath, append: true))
            {
                // Write a header with timestamp
                logWriter.WriteLine($"--- Audit Log: {DateTime.UtcNow:u} ---");
                logWriter.WriteLine($"Applied Localization Region: {region}");
                logWriter.WriteLine();

                // Iterate through all worksheets and their charts
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Obtain chart index via collection since Chart.Index may not be available
                        int chartIndex = sheet.Charts.IndexOf(chart);

                        // Chart name may be empty; provide a fallback identifier
                        string chartName = string.IsNullOrEmpty(chart.Name)
                            ? $"Chart_{chartIndex}"
                            : chart.Name;

                        // Log the chart name together with its worksheet
                        logWriter.WriteLine($"Worksheet: {sheet.Name}, Chart Index: {chartIndex}, Chart Name: {chartName}");
                    }
                }

                logWriter.WriteLine(); // Blank line for readability
            }

            Console.WriteLine("Chart audit information has been logged successfully.");
        }
        catch (Exception ex)
        {
            // Log unexpected errors to console (could be extended to log file)
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
