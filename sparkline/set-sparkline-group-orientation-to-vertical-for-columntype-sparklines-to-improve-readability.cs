using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetSparklineOrientationVertical
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the column‑type sparkline (by column)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(3);
        sheet.Cells["A3"].PutValue(7);
        sheet.Cells["A4"].PutValue(2);
        sheet.Cells["A5"].PutValue(9);

        // Define where the sparkline(s) will be placed (e.g., cell B1)
        CellArea location = new CellArea
        {
            StartRow = 0,   // row 1 (zero‑based)
            EndRow = 0,
            StartColumn = 1, // column B (zero‑based)
            EndColumn = 1
        };

        // Add a column‑type sparkline group with vertical orientation (isVertical = true)
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Column,   // column sparkline
            "A1:A5",                // data range
            true,                   // plot by column (vertical orientation)
            location);              // location range

        // Retrieve the created group (optional, for further customization)
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the sparkline to the group (the Add method of Sparklines collection)
        group.Sparklines.Add(sheet.Name + "!A1:A5", 0, 1);

        // Save the workbook
        workbook.Save("SparklineVerticalOrientation.xlsx");
    }
}