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

        // Populate sample data that the sparkline will represent
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(3);
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(2);

        // Define the cell area where the sparkline will be placed (e.g., cell E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4, // Column E (0‑based index)
            EndColumn = 4
        };

        // Add a sparkline group:
        // - Type: Line
        // - Data range: A1:D1
        // - isVertical: false (plot by column)
        // - Location: the CellArea defined above
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group (optional, demonstrates the Sparklines collection)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Save the workbook as an XLSX file
        workbook.Save("SparklineExample.xlsx", SaveFormat.Xlsx);
    }
}