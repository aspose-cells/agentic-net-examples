// Title: Create a third‑order polynomial trendline for a scatter chart series and read its equation using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add a polynomial trendline of order 3 to a scatter chart | How to get the equation of a polynomial trendline from an Excel chart with Aspose.Cells | Aspose.Cells .NET extract trendline equation after chart calculation | Create or locate a scatter chart and apply a third‑degree trendline using Aspose.Cells | Sample code for adding a polynomial trendline to a chart series in Aspose.Cells C#
// Tags: add polynomial trendline Aspose.Cells C# | scatter chart trendline order 3 Aspose.Cells | extract trendline formula Aspose.Cells .NET | chart calculation Aspose.Cells C# | create scatter chart series Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example checks for an input workbook, creates sample X/Y data if needed, loads the workbook, finds or adds a scatter chart, adds a series referencing the data, inserts a third‑order polynomial trendline, enables equation display, forces chart calculation, reads the generated equation string, prints it, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input workbook exists; if not, create a sample workbook.
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                // Populate sample data for Y (A) and X (B) columns.
                for (int i = 0; i < 10; i++)
                {
                    ws.Cells[i, 0].PutValue(i + 1);          // Y values in column A
                    ws.Cells[i, 1].PutValue((i + 1) * 2);   // X values in column B
                }
                wb.Save(inputPath);
            }

            // Load the existing workbook (file existence already checked above)
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the first scatter chart; if none exists, create one
            Chart scatterChart = null;
            foreach (Chart c in sheet.Charts)
            {
                if (c.Type == ChartType.Scatter)
                {
                    scatterChart = c;
                    break;
                }
            }

            if (scatterChart == null)
            {
                // Add a new scatter chart
                int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 25, 10);
                scatterChart = sheet.Charts[chartIndex];

                // Add a sample series (Y values in column A, X values in column B)
                scatterChart.NSeries.Add("A1:A10", true);
                scatterChart.NSeries[0].XValues = "B1:B10";
            }

            // Ensure the chart has at least one series
            if (scatterChart.NSeries.Count == 0)
                throw new InvalidOperationException("Scatter chart does not contain any series.");

            // Use dynamic to access Trendline APIs that may vary between Aspose.Cells versions
            dynamic series = scatterChart.NSeries[0];
            dynamic trendline = series.Trendlines.Add(TrendlineType.Polynomial);
            trendline.Order = 3;               // Set polynomial order to 3
            trendline.DisplayEquation = true; // Show equation on the chart

            // Force chart calculation so the trendline data becomes available
            scatterChart.Calculate();

            // Retrieve the generated equation string (if supported)
            string equation = trendline.TrendlineEquation ?? "Equation not available";

            // Output the equation
            Console.WriteLine("Polynomial (order 3) trendline equation: " + equation);

            // Save the workbook with the added trendline
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
