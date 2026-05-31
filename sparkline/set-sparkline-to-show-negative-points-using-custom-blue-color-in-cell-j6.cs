using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineNegativePointsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data with negative values (row 6)
            sheet.Cells["A6"].PutValue(5);
            sheet.Cells["B6"].PutValue(-2);
            sheet.Cells["C6"].PutValue(3);
            sheet.Cells["D6"].PutValue(-4);

            // Define the location area for the sparkline (cell J6)
            CellArea location = new CellArea
            {
                StartColumn = 9, // Column J (0‑based index)
                EndColumn = 9,
                StartRow = 5,    // Row 6 (0‑based index)
                EndRow = 5
            };

            // Add a sparkline group of type Line, using the data range A6:D6
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A6:D6", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group (the same range as above)
            group.Sparklines.Add(sheet.Name + "!A6:D6", 5, 9);

            // Enable highlighting of negative points
            group.ShowNegativePoints = true;

            // Set custom blue color for negative points
            CellsColor blueColor = workbook.CreateCellsColor();
            blueColor.Color = Color.Blue;
            group.NegativePointsColor = blueColor;

            // Save the workbook
            workbook.Save("SparklineNegativePointsBlue.xlsx");
        }
    }
}