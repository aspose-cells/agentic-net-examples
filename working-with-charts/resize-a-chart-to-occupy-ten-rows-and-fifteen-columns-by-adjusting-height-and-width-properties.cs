using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartResize
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart (initial position is arbitrary)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Resize the chart to occupy 10 rows (0‑9) and 15 columns (0‑14)
            // Move method defines the upper‑left and lower‑right cell coordinates.
            chart.Move(topRow: 0, leftColumn: 0, bottomRow: 9, rightColumn: 14);

            // Optionally, you can also adjust the pixel size directly via ChartObject if needed:
            // chart.ChartObject.Width = 800;   // example pixel width
            // chart.ChartObject.Height = 400;  // example pixel height

            // Save the workbook
            workbook.Save("ChartResized.xlsx");
        }
    }
}