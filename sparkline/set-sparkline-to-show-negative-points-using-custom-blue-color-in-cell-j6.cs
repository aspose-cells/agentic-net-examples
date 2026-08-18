// Title: Add a line sparkline in cell J6 with blue‑colored negative points using Aspose.Cells for .NET
// Description: Shows how to create a workbook, insert sample data, place a line sparkline in J6, enable ShowNegativePoints, and apply a custom blue color to negative values with Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# sparkline | line sparkline | negative points color | custom blue sparkline | cell J6 | SparklineGroup.ShowNegativePoints | NegativePointsColor | Excel sparkline .NET | financial sparkline visualization
// Common Searches: Aspose.Cells set negative sparkline color | C# line sparkline blue negative points | place sparkline in J6 using Aspose.Cells | ShowNegativePoints property example | customize sparkline negative values color .NET
// Developer Intent: Create a line sparkline at J6 and color its negative points blue.
// Use Cases: Financial statements where declines are highlighted in blue within a sparkline for quick trend analysis. | Sales performance dashboards that emphasize months with negative growth using a blue sparkline marker. | Automated reporting that inserts sparklines and automatically colors loss points for easy visual inspection.
// AI Prompts: Generate C# code with Aspose.Cells to add a line sparkline at J6 and set NegativePointsColor to blue. | Explain the steps to enable ShowNegativePoints and assign a custom color for negative values in a SparklineGroup. | Provide a tutorial for creating a sparkline group, defining its location, and customizing negative point appearance in Aspose.Cells for .NET.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineNegativePointsDemo
{
    // Shows how to create a workbook, insert sample data, place a line sparkline in J6, enable ShowNegativePoints, and apply a custom blue color to negative values with Aspose.Cells in C#.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data with negative values (A1:D1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(-2);
            sheet.Cells["C1"].PutValue(3);
            sheet.Cells["D1"].PutValue(-4);

            // Define the location area for the sparkline (cell J6)
            CellArea location = new CellArea
            {
                StartRow = 5,   // Row 6 (zero‑based index)
                EndRow = 5,
                StartColumn = 9, // Column J (zero‑based index)
                EndColumn = 9
            };

            // Add a sparkline group of type Line using the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group at J6
            group.Sparklines.Add(sheet.Name + "!A1:D1", 5, 9);

            // Enable highlighting of negative points
            group.ShowNegativePoints = true;

            // Set custom blue color for negative points
            CellsColor negativeColor = workbook.CreateCellsColor();
            negativeColor.Color = Color.Blue;
            group.NegativePointsColor = negativeColor;

            // Save the workbook
            workbook.Save("SparklineNegativePointsDemo.xlsx");
        }
    }
}
