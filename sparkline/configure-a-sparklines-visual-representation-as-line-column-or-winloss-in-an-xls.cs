using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineTypeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for three rows (each will use a different sparkline type)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        sheet.Cells["A2"].PutValue(7);
        sheet.Cells["B2"].PutValue(4);
        sheet.Cells["C2"].PutValue(6);
        sheet.Cells["D2"].PutValue(2);

        sheet.Cells["A3"].PutValue(3);
        sheet.Cells["B3"].PutValue(8);
        sheet.Cells["C3"].PutValue(5);
        sheet.Cells["D3"].PutValue(9);

        // Define the location where sparklines will be placed (column E)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 2,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group with an initial type (Line). This will create three sparklines,
        // one for each row of the data range A1:D3.
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D3", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // The group now contains three sparklines (rows 0,1,2). Change the type of each sparkline
        // by setting the group's Type property. We'll demonstrate all three supported types.

        // 1. Set to Line sparkline
        group.Type = SparklineType.Line;
        // Optional: customize appearance for line type
        group.LineWeight = 1.0;
        group.SeriesColor = workbook.CreateCellsColor();
        group.SeriesColor.Color = Color.DarkBlue;

        // 2. Set to Column sparkline (overwrites previous type for the whole group)
        group.Type = SparklineType.Column;
        // Optional: customize appearance for column type
        group.SeriesColor = workbook.CreateCellsColor();
        group.SeriesColor.Color = Color.ForestGreen;

        // 3. Set to Win/Loss sparkline (Stacked enum value)
        group.Type = SparklineType.Stacked;
        // Optional: customize appearance for win/loss type
        group.SeriesColor = workbook.CreateCellsColor();
        group.SeriesColor.Color = Color.Maroon;

        // Save the workbook with the configured sparklines
        workbook.Save("SparklineTypesDemo.xlsx");
    }
}