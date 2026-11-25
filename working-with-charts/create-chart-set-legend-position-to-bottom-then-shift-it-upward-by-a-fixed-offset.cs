// Title: Create a column chart in Aspose.Cells for .NET, set the legend to the bottom, and shift it upward by a fixed offset
// Common Searches: Aspose.Cells C# move chart legend upward after positioning it at the bottom | How to set legend position to bottom and adjust Y offset in an Aspose.Cells chart | Disable automatic legend sizing and change legend coordinates in a .NET workbook | Increase legend width and reposition legend in an Aspose.Cells column chart | Shift chart legend upward by points using Aspose.Cells API
// Tags: column chart legend positioning Aspose.Cells | legend Y coordinate offset Aspose.Cells .NET | disable automatic legend sizing Aspose.Cells | increase legend width Aspose.Cells chart | set legend position bottom Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendShiftDemo
{
    // The example creates a new workbook, adds sample data, generates a column chart, places its legend at the bottom, disables automatic sizing, moves the legend upward by 200 points, widens it, and saves the file as ChartWithShiftedLegend.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Populate some sample data for the chart
            // ------------------------------------------------------------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");

            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // ------------------------------------------------------------
            // Add a column chart to the worksheet
            // ------------------------------------------------------------
            // Parameters: chart type, upper‑left row, upper‑left column,
            // lower‑right row, lower‑right column
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // ------------------------------------------------------------
            // Configure the legend
            // ------------------------------------------------------------
            // 1. Place the legend at the bottom of the chart
            chart.Legend.Position = LegendPositionType.Bottom;

            // 2. Disable automatic sizing so we can move it manually
            chart.Legend.IsAutomaticSize = false;

            // 3. Shift the legend upward by a fixed offset (in points)
            //    The Y property is measured from the top of the chart.
            //    Because the legend is docked at the bottom, we move it
            //    upward by reducing the Y coordinate.
            //    The offset value can be adjusted as needed.
            const int upwardOffset = 200; // points
            chart.Legend.Y = chart.Legend.Y - upwardOffset;

            // Optional: make the legend a little wider so the text does not wrap
            chart.Legend.Width = 3000; // width in points (approx. 42 mm)

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            string outFile = "ChartWithShiftedLegend.xlsx";
            workbook.Save(outFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outFile}");
        }
    }
}
