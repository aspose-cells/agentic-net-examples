// Title: Aspose.Cells .NET: Position Chart Legend at Top‑Right Corner and Retrieve Its Coordinates
// Description: Creates a workbook, adds sample data, inserts a column chart, sets the legend to the top‑right corner using LegendPositionType.Corner, reads the XRatioToChart and YRatioToChart values to confirm placement, outputs the ratios, and saves the file as LegendTopRightCorner.xlsx.
// Keywords: Aspose.Cells chart legend position | LegendPositionType Corner | C# chart legend coordinates | XRatioToChart Aspose.Cells | YRatioToChart Aspose.Cells | .NET Excel chart legend | Aspose.Cells legend placement verification
// Common Searches: Aspose.Cells set legend top right | Get legend XRatioToChart in C# | How to read legend YRatioToChart Aspose.Cells | Chart legend corner position .NET | Verify legend coordinates Aspose.Cells
// Developer Intent: Place a chart legend in the top‑right corner and programmatically confirm its relative X/Y coordinates.
// Use Cases: Standardize legend placement across generated reports. | Automated validation of chart layout in dynamic Excel files. | Adjust surrounding elements based on precise legend positioning.
// AI Prompts: Generate C# code that moves a chart legend to the bottom‑left corner and returns its XRatioToChart and YRatioToChart values using Aspose.Cells. | Explain the relationship between Legend.Position, XRatioToChart, and YRatioToChart in Aspose.Cells for .NET. | Create a unit test that asserts the legend ratios fall within expected ranges after setting Legend.Position to Corner.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendDemo
{
    // Creates a workbook, adds sample data, inserts a column chart, sets the legend to the top‑right corner using LegendPositionType.Corner, reads the XRatioToChart and YRatioToChart values to confirm placement, outputs the ratios, and saves the file as LegendTopRightCorner.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Position the legend at the top‑right corner of the plot area
            chart.Legend.Position = LegendPositionType.Corner;

            // After setting the position, retrieve the relative coordinates
            // (ratio to chart area) to verify where the legend is placed.
            double legendXRatio = chart.Legend.XRatioToChart;
            double legendYRatio = chart.Legend.YRatioToChart;

            Console.WriteLine($"Legend X ratio to chart: {legendXRatio}");
            Console.WriteLine($"Legend Y ratio to chart: {legendYRatio}");

            // Save the workbook (the file will contain the chart with the legend positioned)
            workbook.Save("LegendTopRightCorner.xlsx");
        }
    }
}
