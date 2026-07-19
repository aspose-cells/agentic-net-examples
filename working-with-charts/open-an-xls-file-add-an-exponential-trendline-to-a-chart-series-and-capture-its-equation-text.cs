// Title: Add Exponential Trendline with Equation Display to an Excel Chart using Aspose.Cells for .NET
// Description: Loads an existing XLS workbook, ensures a chart exists (creates a line chart if needed), adds an exponential trendline to the first series, enables the equation label on the chart, and saves the modified file. Demonstrates how to work with charts and trendlines in Aspose.Cells for C#.
// Keywords: Aspose.Cells exponential trendline | C# add trendline to Excel chart | display trendline equation Aspose.Cells | modify existing XLS workbook chart | Aspose.Cells chart creation .NET | trendline equation visibility Excel | Aspose.Cells TrendlineType.Exponential
// Common Searches: how to add exponential trendline to an Excel chart with Aspose.Cells | enable trendline equation display in C# Aspose.Cells | add chart to XLS if none exists Aspose.Cells | retrieve trendline equation text Aspose.Cells .NET | Aspose.Cells add trendline to first series
// Developer Intent: Insert an exponential trendline into the first series of a chart in an existing XLS file and turn on the equation label using Aspose.Cells for .NET.
// Use Cases: Enhance legacy sales reports with a fitted exponential curve and visible equation for quick analysis. | Generate scientific data visualizations that automatically show the exponential fit formula in Excel workbooks. | Automate report preparation in a .NET service where every chart must display its trendline equation when opened.
// AI Prompts: Write C# code with Aspose.Cells that adds an exponential trendline to the first series of a chart and shows the equation on the chart. | Explain why Aspose.Cells does not expose the trendline equation via API and suggest ways to capture the equation text. | Provide a sample that checks for an existing chart, creates a line chart if missing, adds a data series, and configures an exponential trendline with equation visibility.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLS workbook, ensures a chart exists (creates a line chart if needed), adds an exponential trendline to the first series, enables the equation label on the chart, and saves the modified file. Demonstrates how to work with charts and trendlines in Aspose.Cells for C#.
class AddExponentialTrendline
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xls";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart; create one if none exist
            Chart chart;
            if (sheet.Charts.Count > 0)
            {
                chart = sheet.Charts[0];
            }
            else
            {
                // Create a simple line chart (adjust range as needed)
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 8);
                chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("A1:B4", true);
            }

            // Add an exponential trendline to the first series
            int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Exponential);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

            // Enable display of the equation on the chart
            trendline.DisplayEquation = true;

            // Note: Retrieving the equation text directly is not supported via the DataLabels indexer.
            // The equation will be visible on the chart when opened in Excel.
            Console.WriteLine("Exponential trendline added and equation display enabled.");

            // Save the modified workbook
            string outputPath = "output.xls";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
