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
        sheet.Cells["A2"].PutValue(3);
        sheet.Cells["A3"].PutValue(7);
        sheet.Cells["A4"].PutValue(2);
        sheet.Cells["A5"].PutValue(9);

        // Define the cells where the sparklines will be placed (B1:B5)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 4,
            StartColumn = 1,
            EndColumn = 1
        };

        // Add a sparkline group with the data range A1:A5
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:A5", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Optionally, adjust properties of the sparkline(s) if needed
        // Example: set the first sparkline's data range explicitly (not required here)
        // group.Sparklines[0].DataRange = "A1:A5";

        // Save the workbook to an XLSX file
        workbook.Save("SparklineDataRangeDemo.xlsx");
    }
}