using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class DuplicateSparklineGroupExample
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

        // Define the location where the original sparkline will be placed
        CellArea originalLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add the original sparkline group (Line type) to the worksheet
        int originalIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, originalLocation);
        SparklineGroup originalGroup = sheet.SparklineGroups[originalIndex];
        originalGroup.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Duplicate the SparklineGroup by creating a new group with the same settings
        CellArea duplicateLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 5,
            EndColumn = 5
        };
        int duplicateIndex = sheet.SparklineGroups.Add(originalGroup.Type, "A1:D1", false, duplicateLocation);
        SparklineGroup duplicateGroup = sheet.SparklineGroups[duplicateIndex];
        duplicateGroup.Sparklines.Add(sheet.Name + "!A1:D1", 0, 5);

        // Save the workbook to an XLSX file
        workbook.Save("DuplicatedSparklineGroup.xlsx", SaveFormat.Xlsx);
    }
}