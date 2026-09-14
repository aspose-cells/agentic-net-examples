// Title: Add a line sparkline to cell J6 and color its negative points blue with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to insert a line sparkline at J6, enable ShowNegativePoints, and assign a blue CellsColor to NegativePointsColor. | Show how to configure a SparklineGroup to highlight negative values with a custom blue color in a .NET workbook.
// Common Searches: how to highlight negative values in a sparkline using Aspose.Cells C# | set custom blue color for negative points in a line sparkline Aspose.Cells .NET | add sparkline to specific cell J6 with Aspose.Cells example | configure SparklineGroup.ShowNegativePoints and NegativePointsColor in C# | Aspose.Cells create sparkline and customize negative point color
// Tags: line sparkline negative points color Aspose.Cells | sparkline placement J6 Aspose.Cells | custom blue CellsColor for sparkline negatives | SparklineGroup.ShowNegativePoints configuration | Aspose.Cells sparkline styling example

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineNegativePointsDemo
{
    // Demonstrates creating a workbook, adding sample data, inserting a line sparkline at J6, enabling negative point highlighting, applying a custom blue CellsColor to negative points, and saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data containing negative values (A1:D1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(-2);
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["D1"].PutValue(-4);

            // Define the location area for the sparkline (cell J6)
            CellArea location = new CellArea
            {
                StartColumn = 9, // Column J (0‑based index)
                EndColumn = 9,
                StartRow = 5,    // Row 6 (0‑based index)
                EndRow = 5
            };

            // Add a line sparkline group using the data range A1:D1 and place it at J6
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group (required when using Add with location)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 9);

            // Enable highlighting of negative points
            group.ShowNegativePoints = true;

            // Set custom blue color for negative points
            CellsColor blueColor = workbook.CreateCellsColor();
            blueColor.Color = Color.Blue;
            group.NegativePointsColor = blueColor;

            // Save the workbook
            workbook.Save("SparklineNegativePointsDemo.xlsx");
        }
    }
}
