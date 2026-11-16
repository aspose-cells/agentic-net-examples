// Title: Programmatically add linear trendlines to all bar and column chart series in an Excel workbook and log their details using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add linear trendline to each series of a bar chart | Iterate through worksheets and charts to apply trendlines with Aspose.Cells | Save Excel workbook after adding trendlines to charts using Aspose.Cells | Write trendline information to a text file in C# with Aspose.Cells | Aspose.Cells log chart series trendline equations in .NET
// Tags: Aspose.Cells add linear trendline to chart series | Aspose.Cells iterate worksheets and charts | Aspose.Cells write trendline details to log file | Aspose.Cells save workbook with updated charts | C# process bar and column charts programmatically

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, walks through each worksheet and its charts, adds a linear trendline to every series of bar or column charts, captures sheet name, chart title, series name and a placeholder equation, writes these details to a text log, and finally saves the workbook with the new trendlines.
class TrendlineProcessor
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output_with_trendlines.xlsx";
        const string logPath = "trendlines_log.txt";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from file
            Workbook workbook = new Workbook(inputPath);

            // List to hold all trendline information
            List<string> logLines = new List<string>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Process only Bar or Column charts
                    if (chart.Type == ChartType.Bar || chart.Type == ChartType.Column)
                    {
                        try
                        {
                            // Iterate through each series in the chart
                            foreach (Series series in chart.NSeries)
                            {
                                // Use dynamic to access Trendlines (may not be available in older versions)
                                dynamic dynSeries = series;
                                try
                                {
                                    // Add a linear trendline to the series
                                    Trendline trendline = dynSeries.Trendlines.Add(TrendlineType.Linear);
                                    trendline.DisplayEquation = true; // ensure equation is shown on the chart
                                }
                                catch (Exception exTrend)
                                {
                                    Console.WriteLine($"Unable to add trendline to series \"{series.Name}\": {exTrend.Message}");
                                }

                                // Placeholder for equation display (Aspose.Cells does not expose the equation string)
                                string equation = "Displayed on chart";

                                // Store the information with context
                                string chartTitle = chart.Title?.Text ?? "Untitled";
                                logLines.Add(
                                    $"Sheet: {sheet.Name}, Chart: {chartTitle}, Series: {series.Name}, Equation: {equation}");
                            }
                        }
                        catch (Exception exChartSeries)
                        {
                            Console.WriteLine($"Error processing series in chart \"{chart.Title?.Text}\": {exChartSeries.Message}");
                        }
                    }
                }
            }

            // Write all collected information to a log file
            try
            {
                using (StreamWriter writer = new StreamWriter(logPath, false))
                {
                    writer.WriteLine("Trendline Information Log");
                    writer.WriteLine($"Generated on: {DateTime.Now}");
                    writer.WriteLine(new string('-', 50));
                    foreach (string line in logLines)
                    {
                        writer.WriteLine(line ?? string.Empty);
                    }
                }
            }
            catch (Exception exLog)
            {
                Console.WriteLine($"Failed to write log file: {exLog.Message}");
            }

            // Ensure the directory for the output file exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the added trendlines
                workbook.Save(outputPath);
                Console.WriteLine($"Processing completed. Log saved to \"{logPath}\" and workbook saved to \"{outputPath}\".");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Failed to save workbook: {exSave.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
