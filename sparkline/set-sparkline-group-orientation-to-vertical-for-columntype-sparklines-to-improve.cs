using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class SetSparklineOrientationVertical
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the column sparkline (by column)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(3);
        sheet.Cells["A3"].PutValue(7);
        sheet.Cells["A4"].PutValue(2);
        sheet.Cells["A5"].PutValue(9);

        // Define where the sparkline will be placed (single cell)
        CellArea location = new CellArea
        {
            StartRow = 0,   // row 1 (zero‑based)
            EndRow = 0,
            StartColumn = 1, // column B
            EndColumn = 1
        };

        // Add a Column‑type sparkline group, set isVertical = true
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Column,   // sparkline type
            "A1:A5",                // data range
            true,                   // isVertical = true
            location);              // where sparkline is placed

        // Retrieve the created group (optional, for further customization)
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Example: change series color (optional)
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Orange;
        group.SeriesColor = seriesColor;

        // Save the workbook
        workbook.Save("SparklineVerticalOrientation.xlsx");
    }
}