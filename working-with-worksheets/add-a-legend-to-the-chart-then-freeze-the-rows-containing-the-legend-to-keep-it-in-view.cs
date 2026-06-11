using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendFreezeDemo
{
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

            // Add a column chart. The chart occupies rows 5‑20 and columns 0‑8
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the legend is displayed and customize its appearance
            chart.ShowLegend = true;                     // make sure legend is visible
            chart.Legend.Position = LegendPositionType.Bottom; // place legend below the plot area
            chart.Legend.IsAutomaticSize = false;       // allow manual sizing
            chart.Legend.Width = 400;                    // width in pixels
            chart.Legend.Height = 50;                    // height in pixels
            chart.Legend.Font.Size = 12;                 // font size
            chart.Legend.Font.IsBold = true;             // bold font

            // Freeze the rows that contain the chart (including the legend)
            // Freeze up to row 20 (the bottom row of the chart) and column 0
            // This keeps the chart and its legend visible while scrolling
            sheet.FreezePanes(20, 0, 20, 0);

            // Save the workbook
            workbook.Save("ChartWithLegendAndFrozenRows.xlsx");
        }
    }
}