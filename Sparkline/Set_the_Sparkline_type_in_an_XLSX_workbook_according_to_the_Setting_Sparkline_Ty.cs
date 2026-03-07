using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetSparklineType
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell area where the sparkline will be placed (column E)
        CellArea ca = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4, // Column index 4 corresponds to column E
            EndColumn = 4
        };

        // Add a sparkline group with initial type Column
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Column, "A1:D1", false, ca);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Change the sparkline type to Line as required
        group.Type = SparklineType.Line;

        // Add a sparkline to the group (data range A1:D1, placed at row 0, column 4)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Save the workbook to an XLSX file
        workbook.Save("SparklineTypeDemo.xlsx");
    }
}