using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class SparklineTypeDemo
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

        // Define the cell where the first sparkline will be placed (E1)
        CellArea location1 = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group with the default Line type
        int idx1 = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location1);
        SparklineGroup group1 = sheet.SparklineGroups[idx1];
        group1.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Change the sparkline style of the first group to Column
        group1.Type = SparklineType.Column;

        // Define the cell for a second sparkline (E2) and add a Win/Loss (Stacked) type group
        CellArea location2 = new CellArea
        {
            StartRow = 1,
            EndRow = 1,
            StartColumn = 4,
            EndColumn = 4
        };
        int idx2 = sheet.SparklineGroups.Add(SparklineType.Stacked, "A1:D1", false, location2);
        SparklineGroup group2 = sheet.SparklineGroups[idx2];
        group2.Sparklines.Add(sheet.Name + "!A1:D1", 1, 4);
        // Explicitly set the type (already Stacked, but shown for completeness)
        group2.Type = SparklineType.Stacked;

        // Save the workbook to an XLSX file
        workbook.Save("SparklineTypesDemo.xlsx", SaveFormat.Xlsx);
    }
}