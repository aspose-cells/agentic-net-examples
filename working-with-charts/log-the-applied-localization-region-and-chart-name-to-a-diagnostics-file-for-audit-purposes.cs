// Title: Log workbook localization region and chart names to a diagnostics file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells write chart name and workbook culture to a log file | How to append workbook region and chart identifiers to a diagnostics log using Aspose.Cells | Aspose.Cells iterate charts and record localization info in a text file | Save Excel chart audit information with region metadata using Aspose.Cells .NET | Log errors and chart details to a diagnostics file in Aspose.Cells C# example
// Tags: Aspose.Cells chart audit logging C# | write workbook localization info to file | append chart metadata to diagnostics log | iterate Excel worksheets and charts Aspose.Cells | exception handling while logging Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads or creates an Excel workbook, determines the applied localization region (or falls back to the current culture), iterates through each worksheet and its charts, and appends a timestamped entry containing the region and chart name to a diagnostics log file. It also saves the workbook and records any caught exceptions in the same log.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string logFilePath = "diagnostics.log";

        try
        {
            // Ensure the input workbook exists; create a new one if missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // empty workbook
                workbook.Save(inputPath);   // optionally create the file for future runs
            }

            // Retrieve the applied localization region (culture) from workbook settings.
            // If not set, fallback to the current culture.
            string region = workbook.Settings.CultureInfo?.Name ?? CultureInfo.CurrentCulture.Name;

            // Iterate through all worksheets and their charts.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Get the chart's name.
                    string chartName = chart.Name;

                    // Build a log entry with timestamp, region, and chart name.
                    string logEntry = $"{DateTime.UtcNow:u} | Region: {region} | Chart: {chartName}{Environment.NewLine}";

                    // Append the entry to the diagnostics file.
                    File.AppendAllText(logFilePath, logEntry);
                }
            }

            // Save the workbook if any modifications were made (optional).
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log unexpected errors to the diagnostics file.
            string errorEntry = $"{DateTime.UtcNow:u} | ERROR: {ex.Message}{Environment.NewLine}";
            File.AppendAllText(logFilePath, errorEntry);
        }
    }
}
