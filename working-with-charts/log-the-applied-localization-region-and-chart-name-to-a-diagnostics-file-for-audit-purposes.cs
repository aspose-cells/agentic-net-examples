// Title: Log workbook localization region and chart titles to a diagnostics file with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing Excel workbook using Aspose.Cells, sets the workbook's CultureInfo, and appends the culture name and each chart's title to a diagnostics log with timestamps. | Create a diagnostic logger in C# that iterates through all worksheets and charts in a workbook loaded by Aspose.Cells, then writes the applied localization region and chart names to a log file. | Generate a C# example that saves the workbook after logging, includes error handling for missing input files, and ensures the log file is opened in append mode.
// Common Searches: aspnet aspose.cells log workbook culture info and chart names | c# write excel chart titles to a diagnostics log file | how to record localization region of an Excel workbook using Aspose.Cells | append timestamps and chart details to a log while processing Excel with Aspose.Cells | save workbook after logging chart information in C# Aspose.Cells
// Tags: Aspose.Cells log workbook localization | C# write chart titles to file | Aspose.Cells iterate worksheets charts | diagnostics logging with timestamps C# | set workbook CultureInfo Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example checks for the presence of input.xlsx, loads it with Aspose.Cells, sets the workbook's CultureInfo to en-US, opens diagnostics.log in append mode, writes a timestamped entry for the applied localization region, iterates through each worksheet and its charts to log each chart's title (or a placeholder if none), adds a separator line, saves the workbook to output.xlsx, and handles any runtime exceptions.
class DiagnosticsLogger
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string diagnosticsFilePath = "diagnostics.log";

        // Verify input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply a localization region (e.g., English - United States)
            workbook.Settings.CultureInfo = new CultureInfo("en-US");

            // Open the diagnostics file for appending
            using (StreamWriter writer = new StreamWriter(diagnosticsFilePath, true))
            {
                // Log the applied localization region
                writer.WriteLine($"[{DateTime.Now}] Localization Region: {workbook.Settings.CultureInfo.Name}");

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the current worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Retrieve the chart name; if no title is set, use a placeholder
                        string chartName = chart.Title != null && !string.IsNullOrEmpty(chart.Title.Text)
                            ? chart.Title.Text
                            : "Unnamed Chart";

                        // Log the chart name
                        writer.WriteLine($"[{DateTime.Now}] Chart Name: {chartName}");
                    }
                }

                // Add a separator for readability
                writer.WriteLine(new string('-', 50));
            }

            // Save the workbook after any modifications
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
