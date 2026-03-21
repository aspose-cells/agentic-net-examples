using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (A1:D1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["D1"].PutValue(2);

            // Define the cell where the sparkline will be placed (E1)
            CellArea location = CellArea.CreateCellArea("E1", "E1");

            // Add a sparkline group of type Line with the data range and location
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,          // Sparkline type
                "A1:D1",                     // Data range
                false,                       // Plot by row (horizontal)
                location);                   // Location range

            // Retrieve the created sparkline group
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add a sparkline to the group (data range, row index, column index)
            // Row and column are zero‑based; E1 corresponds to row 0, column 4
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // Optional: customize appearance
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            group.SeriesColor = seriesColor;

            group.ShowHighPoint = true;
            group.ShowLowPoint = true;

            CellsColor highColor = workbook.CreateCellsColor();
            highColor.Color = Color.Green;
            group.HighPointColor = highColor;

            CellsColor lowColor = workbook.CreateCellsColor();
            lowColor.Color = Color.Red;
            group.LowPointColor = lowColor;

            group.LineWeight = 1.0;

            // Save the workbook as an XLSX file
            workbook.Save("SparklineDemo.xlsx");
        }
    }
}