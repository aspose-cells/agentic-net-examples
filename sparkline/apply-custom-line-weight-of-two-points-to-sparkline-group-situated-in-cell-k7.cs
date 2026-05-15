using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ApplySparklineLineWeight
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for the sparkline (adjust as needed)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the location where the sparkline will be placed (cell K7)
        // Column K -> index 10, Row 7 -> index 6 (zero‑based)
        CellArea location = new CellArea
        {
            StartColumn = 10,
            EndColumn = 10,
            StartRow = 6,
            EndRow = 6
        };

        // Add a sparkline group of type Line, using the sample data range
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the sparkline to the group (required for the group to exist)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 10);

        // Apply a custom line weight of 2 points
        group.LineWeight = 2.0;

        // Save the workbook
        workbook.Save("SparklineWithCustomLineWeight.xlsx", SaveFormat.Xlsx);
    }
}