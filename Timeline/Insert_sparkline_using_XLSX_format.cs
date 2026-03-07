using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

class InsertSparklineExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that the sparkline will represent
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell area where the sparkline will be placed (E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group:
        // - Type: Line
        // - Data range: A1:D1
        // - isVertical: false (plot by columns)
        // - Location: the CellArea defined above
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group at row 0, column 4 (cell E1)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Save the workbook in XLSX format
        workbook.Save("SparklineInserted.xlsx", SaveFormat.Xlsx);
    }
}