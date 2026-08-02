using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartPositionDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Fruits");
            worksheet.Cells["A3"].PutValue("Vegetables");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(50);
            worksheet.Cells["B3"].PutValue(30);

            // Add a column chart. Initial position is rows 5‑15, columns 0‑5
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Align the chart's top‑left corner to cell B2 (row 1, column 1)
            // and bottom‑right corner to cell I12 (row 11, column 8)
            // Note: Row and column indices are zero‑based.
            chart.Move(topRow: 1, leftColumn: 1, bottomRow: 11, rightColumn: 8);

            // Optional: make the chart move and size with the cells
            chart.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            workbook.Save("ChartAlignedWithCells.xlsx");
        }
    }
}