using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class AssignNewDataRangeToDuplicatedSparkline
{
    static void Main()
    {
        // Load or create a workbook (replace "input.xlsx" with your file if needed)
        Workbook workbook = new Workbook(); // creates a new workbook
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the original sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Populate sample data for the new data range
        sheet.Cells["A2"].PutValue(7);
        sheet.Cells["B2"].PutValue(4);
        sheet.Cells["C2"].PutValue(6);
        sheet.Cells["D2"].PutValue(2);

        // Define the location where the first sparkline will be placed (cell E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group with an initial data range (A1:D1)
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the original sparkline to the group (already created by Add above)
        // The Add method of SparklineGroupCollection creates the sparkline, so we just retrieve it
        Sparkline originalSparkline = group.Sparklines[0];

        // Duplicate the sparkline: add a new sparkline at a different location (cell E2)
        int duplicateIndex = group.Sparklines.Add(originalSparkline.DataRange, 1, 4);
        Sparkline duplicatedSparkline = group.Sparklines[duplicateIndex];

        // Assign a new data range to the duplicated sparkline (A2:D2)
        duplicatedSparkline.DataRange = "A2:D2";

        // Save the workbook
        workbook.Save("DuplicatedSparklineWithNewRange.xlsx");
    }
}