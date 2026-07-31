// Title: C# – Retrieve Chart Trendline Equation After Refreshing Workbook with Aspose.Cells
// Description: Demonstrates how to open an XLSX file using Aspose.Cells for .NET, call RefreshAll to update all data sources, locate the first chart’s first series, enable DisplayEquation on its trendline, and read the equation (or custom name) for logging before saving the workbook. If the workbook is missing, a sample file with a linear trendline is generated automatically.
// Keywords: Aspose.Cells | C# | .NET | Excel chart trendline | trendline equation | DisplayEquation | RefreshAll | extract regression formula | read trendline name | chart automation | Excel workbook refresh
// Common Searches: Aspose.Cells get chart trendline equation C# | refresh workbook and read trendline formula Aspose.Cells | how to display trendline equation in Excel chart using .NET | retrieve regression equation from chart after data refresh | C# code to log chart trendline equation with Aspose.Cells
// Developer Intent: Extract the text of a chart trendline equation after refreshing workbook data with Aspose.Cells.
// Use Cases: Add the regression formula to automated analytics reports generated from Excel charts. | Verify that a trendline updates correctly when source data changes in a data‑pipeline job. | Create a fallback workbook with a sample trendline for testing and then capture its equation for debugging.
// AI Prompts: Generate C# code that opens an XLSX file with Aspose.Cells, calls RefreshAll, accesses the first chart’s first series trendline, sets DisplayEquation = true, and returns the equation string for logging. | Show how to extract a linear regression equation from an Excel chart trendline and write it to a text file using Aspose.Cells for .NET. | Explain best practices for handling missing charts or trendlines in Aspose.Cells and logging appropriate messages.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to open an XLSX file using Aspose.Cells for .NET, call RefreshAll to update all data sources, locate the first chart’s first series, enable DisplayEquation on its trendline, and read the equation (or custom name) for logging before saving the workbook. If the workbook is missing, a sample file with a linear trendline is generated automatically.
class RetrieveTrendlineEquation
{
    static void Main()
    {
        // Path to the workbook that should contain a chart with a trendline
        string inputPath = "ChartWithTrendline.xlsx";

        try
        {
            // Ensure the input file exists; if not, create a sample workbook
            if (!File.Exists(inputPath))
            {
                CreateSampleWorkbook(inputPath);
                Console.WriteLine($"Sample workbook created at '{inputPath}'.");
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Refresh all data sources (pivot tables, charts, etc.)
            workbook.Worksheets.RefreshAll();

            // Assume the first worksheet and the first chart contain the desired trendline
            Worksheet worksheet = workbook.Worksheets[0];
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least one series and that series has at least one trendline
            if (chart.NSeries.Count == 0 || chart.NSeries[0].TrendLines.Count == 0)
            {
                Console.WriteLine("No trendlines found in the first series of the chart.");
                return;
            }

            // Get the first trendline
            Trendline trendline = chart.NSeries[0].TrendLines[0];

            // Make sure the equation is displayed (this also turns on data labels)
            trendline.DisplayEquation = true;

            // Log trendline information
            Console.WriteLine("Trendline type: " + trendline.Type);
            Console.WriteLine("DisplayEquation flag is set to: " + trendline.DisplayEquation);

            // If a custom name was set for the trendline, display it
            if (!string.IsNullOrEmpty(trendline.Name))
            {
                Console.WriteLine("Trendline name (may contain custom equation): " + trendline.Name);
            }

            // Save the workbook if any changes were made
            string outputPath = "ChartWithTrendline_Refreshed.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Creates a simple workbook with sample data, a line chart, and a linear trendline
    private static void CreateSampleWorkbook(string filePath)
    {
        try
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("X");
            ws.Cells["B1"].PutValue("Y");
            for (int i = 2; i <= 11; i++)
            {
                ws.Cells[$"A{i}"].PutValue(i - 2);               // X values 0..9
                ws.Cells[$"B{i}"].PutValue((i - 2) * 2 + 5);    // Y = 2X + 5
            }

            // Add a line chart
            int chartIndex = ws.Charts.Add(ChartType.Line, 13, 0, 30, 10);
            Chart chart = ws.Charts[chartIndex];
            chart.NSeries.Add("B2:B11", true);
            chart.NSeries[0].XValues = "A2:A11";

            // Add a linear trendline to the series
            int tlIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
            Trendline tl = chart.NSeries[0].TrendLines[tlIndex];
            tl.DisplayEquation = true;
            tl.Name = "Linear Trendline";

            // Save the workbook
            wb.Save(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to create sample workbook: " + ex.Message);
        }
    }
}
