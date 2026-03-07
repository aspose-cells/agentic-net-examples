using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineGroupExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (row 0, columns A-D)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(3);
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(2);

        // Define the location where the sparkline will be placed (cell E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4, // Column E (0‑based index)
            EndColumn = 4
        };

        // Add a sparkline group of type Line, using the data range A1:D1,
        // plotting horizontally (isVertical = false), and placing the sparkline in E1
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,
            "A1:D1",
            false,
            location);

        // Retrieve the created SparklineGroup
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group (data range and location are already set by Add,
        // but we can add additional sparklines if needed)
        // Here we add the same range to the same cell for demonstration
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Optional: customize appearance (e.g., show high/low points)
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;
        group.HighPointColor.Color = System.Drawing.Color.Green;
        group.LowPointColor.Color = System.Drawing.Color.Red;
        group.LineWeight = 1.0;

        // Save the workbook to an XLSX file
        workbook.Save("SparklineGroupExample.xlsx", SaveFormat.Xlsx);
    }
}