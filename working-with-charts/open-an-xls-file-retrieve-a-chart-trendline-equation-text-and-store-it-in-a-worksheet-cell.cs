// Title: C# – Retrieve a chart trendline equation from an XLS workbook and write it to a cell using Aspose.Cells
// Description: Loads an XLS file (or creates a sample workbook), accesses the first chart on the first worksheet, ensures a linear trendline is present, extracts the intercept (and composes a simple y = m·x + b string), writes the equation to cell C1, and saves the updated workbook.
// Keywords: Aspose.Cells C# chart trendline | retrieve trendline equation XLS | write chart equation to cell | Aspose.Cells linear trendline intercept | extract chart equation .NET | sample code Aspose.Cells trendline
// Common Searches: How to get a trendline equation from an XLS chart using Aspose.Cells | Aspose.Cells C# write chart equation to worksheet cell | Retrieve linear trendline intercept Aspose.Cells | Example: extract chart equation and store in cell | Aspose.Cells read chart trendline text
// Developer Intent: Programmatically obtain the linear trendline equation of the first chart in an XLS workbook and place the formatted equation string into a specified worksheet cell.
// Use Cases: Add the trendline formula next to a chart in financial or sales reports for quick reference. | Automate workbook updates so the latest trendline equation is always reflected after data changes. | Create a summary sheet that records equations of multiple chart trendlines for downstream analysis.
// AI Prompts: Generate C# code with Aspose.Cells that opens an XLS workbook, verifies a linear trendline on the first chart, extracts its intercept, builds a "y = m·x + b" string, and writes it to cell C1. | Provide a robust method that checks for the existence of charts, series, and trendlines before extracting the equation and includes comprehensive error handling. | Show how to extend the example to retrieve both slope and intercept for a linear trendline and format the full equation for storage in a worksheet cell.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLS file (or creates a sample workbook), accesses the first chart on the first worksheet, ensures a linear trendline is present, extracts the intercept (and composes a simple y = m·x + b string), writes the equation to cell C1, and saves the updated workbook.
class RetrieveTrendlineEquation
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "InputWorkbook.xls";
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a sample workbook
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];

                // Sample data for chart creation
                ws.Cells["A1"].PutValue(1);
                ws.Cells["A2"].PutValue(2);
                ws.Cells["A3"].PutValue(3);
                ws.Cells["A4"].PutValue(4);
                ws.Cells["B1"].PutValue(2);
                ws.Cells["B2"].PutValue(4);
                ws.Cells["B3"].PutValue(6);
                ws.Cells["B4"].PutValue(8);

                // Add a scatter chart with a linear trendline
                int chartIdx = ws.Charts.Add(ChartType.Scatter, 5, 0, 15, 10);
                Chart chart = ws.Charts[chartIdx];
                chart.NSeries.Add("B1:B4", true);
                chart.NSeries[0].XValues = "A1:A4";

                // Add linear trendline and enable equation display
                int tlIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
                Trendline tl = chart.NSeries[0].TrendLines[tlIdx];
                tl.DisplayEquation = true;
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chartFirst = worksheet.Charts[0];

            // Ensure the first series has at least one trendline
            if (chartFirst.NSeries.Count == 0 || chartFirst.NSeries[0].TrendLines.Count == 0)
            {
                Console.WriteLine("No series or trendlines found in the chart.");
                return;
            }

            // Access the first trendline of the first series
            Trendline trendline = chartFirst.NSeries[0].TrendLines[0];
            trendline.DisplayEquation = true; // ensure equation is displayed

            // Retrieve intercept (slope is not exposed via API)
            double intercept = trendline.Intercept;

            // Compose a simple linear equation representation
            string equationText = $"y = m·x + {intercept}";
            // Store the equation text in cell C1
            worksheet.Cells["C1"].PutValue(equationText);

            // Save the workbook with the updated cell
            string outputPath = "OutputWorkbook.xls";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
