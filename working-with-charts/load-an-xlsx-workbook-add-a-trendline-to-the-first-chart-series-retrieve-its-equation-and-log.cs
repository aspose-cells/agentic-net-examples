// Title: C# – Add a Linear Trendline and Show Equation on the First Chart Series with Aspose.Cells
// Description: Loads an existing XLSX file, checks for a chart on the first worksheet, adds a linear trendline to the chart's first series, enables equation display, logs each action, and saves the updated workbook to a new file.
// Keywords: Aspose.Cells trendline C# | add linear trendline Aspose.Cells | display chart equation .NET | retrieve trendline equation Aspose.Cells | save workbook after chart edit
// Common Searches: Aspose.Cells add linear trendline to chart series | show trendline equation in Excel chart using C# | how to modify chart and save workbook with Aspose.Cells | C# code for adding trendline to existing Excel file
// Developer Intent: Insert a linear trendline into the first series of an existing chart and make its equation visible.
// Use Cases: Add forecasting lines to sales dashboards automatically. | Annotate financial performance charts with equations for report generation. | Batch‑process workbooks to enrich primary charts before distribution.
// AI Prompts: Generate C# code that opens an XLSX workbook, adds a polynomial trendline to the second series of the first chart, displays the equation, and saves the file using Aspose.Cells. | Explain how to extract the equation string of a trendline after setting DisplayEquation = true with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLSX file, checks for a chart on the first worksheet, adds a linear trendline to the chart's first series, enables equation display, logs each action, and saves the updated workbook to a new file.
class TrendlineExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart
            Chart chart = worksheet.Charts[0];

            // Add a linear trendline to the first series of the chart
            int trendlineIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIndex];

            // Make the equation visible on the chart
            trendline.DisplayEquation = true;

            // Inform the user that the trendline has been added
            Console.WriteLine("Linear trendline added and equation set to display on the chart.");

            // Save the workbook with the added trendline
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
