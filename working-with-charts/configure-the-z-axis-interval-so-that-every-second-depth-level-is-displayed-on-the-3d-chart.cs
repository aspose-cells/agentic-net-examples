// Title: Aspose.Cells .NET – Set Z‑Axis interval to show every other depth level in a 3‑D column chart
// Description: Creates a workbook, adds sample data, inserts a 3‑D column chart, sets DepthPercent to 150 % and uses CategoryAxis.TickMarkSpacing = 2 so that only alternate depth labels appear. The chart is calculated to obtain wall depth in pixels and then saved.
// Keywords: Aspose.Cells Z axis interval | 3D column chart tick mark spacing | display alternate depth levels | CategoryAxis.TickMarkSpacing .NET | retrieve chart wall depth pixels | Aspose.Cells 3D chart customization
// Common Searches: Aspose.Cells show every other depth label in 3D chart | set Z‑axis interval Aspose.Cells C# | tick mark spacing for 3D column chart Aspose | how to get wall depth in pixels Aspose.Cells | adjust depth percent and Z‑axis labeling .NET
// Developer Intent: Show only alternate depth labels on the Z‑axis of a 3‑D column chart.
// Use Cases: Simplify a multi‑category sales chart by displaying depth labels for every second category. | Create a financial dashboard where 3‑D column charts have reduced label clutter for clearer analysis. | Programmatically measure chart wall depth after layout changes to fine‑tune report formatting.
// AI Prompts: Write C# code with Aspose.Cells that configures a 3‑D column chart to display every second depth level and returns the wall depth in pixels. | Explain how CategoryAxis.TickMarkSpacing controls Z‑axis labeling in Aspose.Cells 3‑D charts and suggest alternative techniques for interval control. | Generate a complete Aspose.Cells example that adds sample data, creates a 3‑D column chart, sets DepthPercent to 150 %, configures TickMarkSpacing to 2, calculates the chart, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Creates a workbook, adds sample data, inserts a 3‑D column chart, sets DepthPercent to 150 % and uses CategoryAxis.TickMarkSpacing = 2 so that only alternate depth labels appear. The chart is calculated to obtain wall depth in pixels and then saved.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the 3‑D chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Q1");
                worksheet.Cells["A3"].PutValue("Q2");
                worksheet.Cells["A4"].PutValue("Q3");
                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(150);
                worksheet.Cells["B4"].PutValue(180);
                worksheet.Cells["C1"].PutValue("Series2");
                worksheet.Cells["C2"].PutValue(90);
                worksheet.Cells["C3"].PutValue(110);
                worksheet.Cells["C4"].PutValue(130);

                // Add a 3‑D column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure the depth of the chart (Z‑axis size)
                chart.DepthPercent = 150; // 150 % depth

                // Adjust tick mark spacing to display every second depth level
                chart.CategoryAxis.TickMarkSpacing = 2;

                // Calculate the chart to obtain pixel‑based depth information
                chart.Calculate();

                // Retrieve pixel depth after calculation
                int depthInPixels = chart.Walls.DepthPx;
                Console.WriteLine("Depth in pixels after calculation: " + depthInPixels);

                // Save the workbook
                string outputPath = "ZAxisIntervalEverySecondDepthLevel.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
