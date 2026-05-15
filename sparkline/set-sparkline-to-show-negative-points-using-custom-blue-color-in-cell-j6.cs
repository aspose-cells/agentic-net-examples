using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data with negative values in row 6 (A6:D6)
        sheet.Cells["A6"].PutValue(5);
        sheet.Cells["B6"].PutValue(-2);
        sheet.Cells["C6"].PutValue(3);
        sheet.Cells["D6"].PutValue(-4);

        // Define the cell area where the sparkline will be placed (J6)
        CellArea sparklineLocation = new CellArea
        {
            StartRow = 5,   // Row 6 (0‑based index)
            EndRow = 5,
            StartColumn = 9, // Column J (0‑based index)
            EndColumn = 9
        };

        // Add a line sparkline group that uses the data range A6:D6 and places the sparkline in J6
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A6:D6", false, sparklineLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the sparkline to the group (optional, but ensures the sparkline is created)
        group.Sparklines.Add(sheet.Name + "!A6:D6", 5, 9);

        // Enable highlighting of negative points
        group.ShowNegativePoints = true;

        // Set a custom blue color for the negative points
        CellsColor blueColor = workbook.CreateCellsColor();
        blueColor.Color = Color.Blue;
        group.NegativePointsColor = blueColor;

        // Save the workbook
        workbook.Save("SparklineNegativePointsBlue.xlsx");
    }
}