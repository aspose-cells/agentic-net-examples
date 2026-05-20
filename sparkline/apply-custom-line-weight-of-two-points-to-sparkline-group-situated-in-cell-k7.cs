using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for the sparkline (row 1, columns A to D)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the location cell K7 (column K = 10, row 7 = index 6)
        CellArea location = new CellArea
        {
            StartRow = 6,
            EndRow = 6,
            StartColumn = 10,
            EndColumn = 10
        };

        // Add a line sparkline group with the data range A1:D1 and place it at K7
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // (Optional) Explicitly add the sparkline to the group
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 10);

        // Apply a custom line weight of two points
        group.LineWeight = 2.0;

        // Save the workbook
        workbook.Save("SparklineLineWeight_K7.xlsx", SaveFormat.Xlsx);
    }
}