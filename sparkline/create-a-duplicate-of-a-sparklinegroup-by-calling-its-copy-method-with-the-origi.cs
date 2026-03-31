using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class DuplicateSparklineGroup
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data for the sparkline
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

        // Add the original sparkline group (Line type) with data range A1:D1
        int originalIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, originalLocation);
        SparklineGroup originalGroup = sheet.SparklineGroups[originalIndex];

        // Add a sparkline to the original group
        originalGroup.Sparklines.Add("A1:D1", 0, 4);

        // -------------------------------------------------
        // Duplicate the sparkline group by recreating it
        // -------------------------------------------------

        // Define new location for the duplicate sparkline (e.g., column F)
        CellArea newLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 5,
            EndColumn = 5
        };

        // Create a new sparkline group with the same settings as the original
        int duplicateIndex = sheet.SparklineGroups.Add(originalGroup.Type, "A1:D1", false, newLocation);
        SparklineGroup duplicateGroup = sheet.SparklineGroups[duplicateIndex];

        // Copy display hidden setting if needed
        duplicateGroup.DisplayHidden = originalGroup.DisplayHidden;

        // Copy sparklines from the original group
        foreach (Sparkline sp in originalGroup.Sparklines)
        {
            duplicateGroup.Sparklines.Add(sp.DataRange, sp.Row, sp.Column);
        }

        // Save the workbook to an XLSX file
        workbook.Save("DuplicateSparklineGroup.xlsx", SaveFormat.Xlsx);
    }
}