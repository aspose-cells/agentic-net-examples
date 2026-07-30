// Title: Add Linear Trendlines to All Bar & Column Charts in Excel with Aspose.Cells (C#)
// Description: Loads an existing workbook, scans every worksheet for bar or column charts, adds a linear trendline to each series, shows the equation and R‑squared value, logs chart details to a text file, and saves the updated file.
// Keywords: Aspose.Cells add trendline C# | linear trendline Excel chart Aspose | log trendline equation R squared | iterate charts workbook Aspose.Cells | save workbook with trendlines .NET | export chart trendline details text file | C# chart automation Aspose.Cells
// Common Searches: how to add linear trendlines to all bar charts using Aspose.Cells | Aspose.Cells C# log trendline equation and R squared | iterate through charts in an Excel workbook Aspose | save workbook after adding trendlines Aspose.Cells | export trendline details to a text file C#
// Developer Intent: Automatically attach linear trendlines to every series of bar/column charts in a workbook and record their equations and R‑squared values in a log file.
// Use Cases: Generate forecast lines for sales bar charts and capture equations for reporting. | Batch‑process multiple worksheets to add trendlines and produce an audit log. | Prepare presentation‑ready Excel files with visible trendline equations on all column charts.
// AI Prompts: Create C# code that adds exponential trendlines to line chart series with Aspose.Cells and writes coefficients to a CSV file. | Refactor the sample to output trendline equations and R‑squared values in JSON instead of plain text. | Show how to modify the program so only the first series of each chart receives a moving‑average trendline.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;
using System.Text;

// Loads an existing workbook, scans every worksheet for bar or column charts, adds a linear trendline to each series, shows the equation and R‑squared value, logs chart details to a text file, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";
            const string logPath = "TrendlineLog.txt";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // StringBuilder to collect log information
            StringBuilder logBuilder = new StringBuilder();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet using index for identification
                for (int chartIdx = 0; chartIdx < sheet.Charts.Count; chartIdx++)
                {
                    Chart chart = sheet.Charts[chartIdx];

                    // Process only column (bar) charts
                    if (chart.Type == ChartType.Column || chart.Type == ChartType.Bar)
                    {
                        // Iterate through each series in the chart
                        for (int s = 0; s < chart.NSeries.Count; s++)
                        {
                            // Add a linear trendline to the current series
                            int trendlineIndex = chart.NSeries[s].TrendLines.Add(TrendlineType.Linear);
                            Trendline trendline = chart.NSeries[s].TrendLines[trendlineIndex];

                            // Configure the trendline to display equation and R‑squared value
                            trendline.DisplayEquation = true;
                            trendline.DisplayRSquared = true;
                            trendline.Name = $"Series{s + 1} Linear";

                            // Log details about the added trendline
                            logBuilder.AppendLine($"Worksheet: {sheet.Name}, Chart Index: {chartIdx}, Series: {s}");
                            logBuilder.AppendLine($"  Trendline Type: {trendline.Type}");
                            logBuilder.AppendLine($"  Equation displayed on chart: Yes");
                            logBuilder.AppendLine($"  R‑Squared displayed on chart: Yes");
                        }
                    }
                }
            }

            // Write the collected information to a log file
            File.WriteAllText(logPath, logBuilder.ToString());

            // Save the modified workbook
            workbook.Save(outputPath);

            Console.WriteLine("Processing completed successfully.");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
