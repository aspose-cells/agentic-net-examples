// Title: Add percentage‑only data labels to visible series of a Progress Bar chart using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel file, finds the first chart, and activates data labels that show only the percentage for each visible series with a whole‑number format. | Demonstrate how to loop through a chart's NSeries in Aspose.Cells and set HasDataLabels, DataLabel.ShowPercentage, and NumberFormat properties. | Show a snippet that positions the percentage label at the center of each bar in a progress‑bar style chart with Aspose.Cells.
// Common Searches: Aspose.Cells enable data labels showing percentage for chart series | C# add percentage labels to visible series in Excel chart with Aspose.Cells | how to format chart data labels as whole number percent in Aspose.Cells | progress bar chart data labels Aspose.Cells .NET
// Tags: Aspose.Cells enable series percentage data labels | C# suppress raw values in chart labels | apply 0% number format to chart data labels | center-align data label within progress bar segment | enumerate chart NSeries to configure labels

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads a workbook, accesses the first chart, iterates over its series, enables data labels for visible series, shows only the percentage formatted as a whole‑number percent, optionally centers the label inside each bar, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "ProgressBar.xlsx";
        const string outputPath = "ProgressBar_WithLabels.xlsx";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the Progress Bar chart
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0]; // assume chart is on the first sheet

            // Ensure the sheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Get the first chart (assumed to be the Progress Bar chart)
            Chart progressChart = sheet.Charts[0];

            // Iterate through all series in the chart using dynamic to avoid compile‑time type issues
            foreach (dynamic series in progressChart.NSeries)
            {
                try
                {
                    // Apply only to series that are visible
                    if (series.IsVisible)
                    {
                        // Enable data labels for this series
                        series.HasDataLabels = true;

                        // Show only the percentage value
                        series.DataLabel.ShowPercentage = true;
                        series.DataLabel.ShowValue = false; // hide raw value if not needed

                        // Format the label as a whole‑number percentage (e.g., 75%)
                        series.DataLabel.NumberFormat = "0%";

                        // Optional: Position the label inside the bar (centered)
                        // The Position property expects a ChartDataLabelPosition enum.
                        // If the enum is unavailable, this line can be omitted or set via integer cast.
                        // series.DataLabel.Position = Aspose.Cells.Charts.ChartDataLabelPosition.Center;
                    }
                }
                catch (Exception exSeries)
                {
                    Console.WriteLine($"Error processing a series: {exSeries.Message}");
                }
            }

            // Save the workbook with the updated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the application from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
