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
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (row 1, columns A‑D)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed (cell E1)
            CellArea location = new CellArea
            {
                StartRow = 0,   // Row index 0 (A1‑based)
                EndRow = 0,
                StartColumn = 4, // Column index 4 => column E
                EndColumn = 4
            };

            // Add a sparkline group of type Line, using the data range A1:D1,
            // not vertical (plot by row), and place the sparkline in the defined location
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add a sparkline to the group (data range A1:D1, placed at row 0, column 4 => E1)
            group.Sparklines.Add("A1:D1", 0, 4);

            // Optional: customize appearance
            // Set the series (line) color to Orange
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            group.SeriesColor = seriesColor;

            // Highlight high and low points
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;

            // Set colors for high and low points
            CellsColor highColor = workbook.CreateCellsColor();
            highColor.Color = Color.Green;
            group.HighPointColor = highColor;

            CellsColor lowColor = workbook.CreateCellsColor();
            lowColor.Color = Color.Red;
            group.LowPointColor = lowColor;

            // Save the workbook as an XLSX file
            workbook.Save("SparklinesCreated.xlsx", SaveFormat.Xlsx);
        }
    }
}