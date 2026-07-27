using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartResize
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data for the chart (optional, but chart needs data)
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart; initial position is temporary
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Position the chart so that its upper‑left corner is at cell D5 (row 4, column 3)
            // Use the Move method to anchor the chart; the size will be defined by WidthPt/HeightPt
            chart.Move(4, 3, 4, 3);   // topRow, leftColumn, bottomRow, rightColumn

            // Resize the chart: width = 500 points, height = 300 points
            chart.ChartObject.WidthPt = 500;   // Width in points
            chart.ChartObject.HeightPt = 300;  // Height in points

            // Save the workbook
            workbook.Save("ResizedChart.xlsx", SaveFormat.Xlsx);
        }
    }
}