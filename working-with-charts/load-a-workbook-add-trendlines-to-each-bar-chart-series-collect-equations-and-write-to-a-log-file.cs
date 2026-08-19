// Title: Add Linear Trendlines to All Bar and Column Charts in a Workbook with Aspose.Cells for .NET and Log the Results
// Description: Loads an existing Excel workbook, scans every worksheet for bar or column charts, adds a linear trendline with equation and R‑squared display to each series, writes chart, series and trendline details to a log file, and saves the updated workbook.
// Keywords: Aspose.Cells | .NET | C# | add trendline to chart | bar chart trendline | column chart trendline | trendline equation | R-squared display | log trendline details | Excel workbook automation | batch chart processing
// Common Searches: how to add linear trendlines to bar charts using Aspose.Cells C# | Aspose.Cells log trendline equation to text file | iterate all charts in a workbook and add trendlines .NET | save workbook after modifying chart series Aspose.Cells | C# code for adding trendlines to Excel charts
// Developer Intent: Automatically add linear trendlines (with equation and R‑squared) to every series of bar and column charts in an Excel workbook and record the chart, series and trendline information in a log file.
// Use Cases: Enhance monthly sales dashboards by inserting trendlines into bar charts and keeping an audit trail of the equations. | Create a nightly batch job that processes multiple workbooks, adds trendlines to engineering data charts, and generates a detailed log for downstream analysis. | Provide a reproducible method for financial analysts to enrich Excel reports with trendlines while capturing metadata for compliance reporting.
// AI Prompts: Generate C# code with Aspose.Cells that adds exponential trendlines to line charts and exports the equations to a JSON file. | Refactor the given trendline example to write the logged information into a CSV file with columns for worksheet, chart index, series, trendline type, and equation. | Explain how to customize trendline appearance (color, dash style, thickness) for each series using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing Excel workbook, scans every worksheet for bar or column charts, adds a linear trendline with equation and R‑squared display to each series, writes chart, series and trendline details to a log file, and saves the updated workbook.
class TrendlineProcessor
{
    static void Main()
    {
        // Paths for input workbook, output workbook and log file
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";
        string logPath = "trendlines.log";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Clear previous log content
            File.WriteAllText(logPath, string.Empty);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Process only column or bar charts
                    if (chart.Type == ChartType.Column || chart.Type == ChartType.Bar)
                    {
                        // Determine chart index for logging (Aspose.Cells Chart has no Index property)
                        int chartIdx = sheet.Charts.IndexOf(chart);

                        // Iterate through each series in the chart
                        for (int seriesIdx = 0; seriesIdx < chart.NSeries.Count; seriesIdx++)
                        {
                            // Add a linear trendline to the current series
                            int trendlineIdx = chart.NSeries[seriesIdx].TrendLines.Add(TrendlineType.Linear);
                            Trendline trendline = chart.NSeries[seriesIdx].TrendLines[trendlineIdx];

                            // Configure trendline to display equation and R‑squared value
                            trendline.DisplayEquation = true;
                            trendline.DisplayRSquared = true;

                            // Assign a custom name for easier identification
                            trendline.Name = $"Series{seriesIdx + 1} Linear Trendline";

                            // Log details of the added trendline
                            string logEntry = $"Worksheet: {sheet.Name}, Chart Index: {chartIdx}, Series: {seriesIdx}, " +
                                              $"Trendline Type: {trendline.Type}, Name: {trendline.Name}";
                            File.AppendAllText(logPath, logEntry + Environment.NewLine);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Processing completed. Output saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
            File.AppendAllText(logPath, $"Error: {ex}{Environment.NewLine}");
        }
    }
}
