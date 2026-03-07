using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (A1:D1)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(3);
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(2);

        // Define where the sparkline will be placed (cell E1)
        CellArea location = CellArea.CreateCellArea("E1", "E1");

        // Add a sparkline group of type Line with the data range and location
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group (data range, row index, column index)
        // Row 0 (first row), Column 4 (E column) matches the location defined above
        group.Sparklines.Add("A1:D1", 0, 4);

        // Save the workbook as an XLSX file
        workbook.Save("SparklineDemo.xlsx", SaveFormat.Xlsx);
    }
}