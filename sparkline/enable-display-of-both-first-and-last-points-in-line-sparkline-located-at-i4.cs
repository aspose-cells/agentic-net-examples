// Title: Aspose.Cells for .NET – Show First & Last Points of a Line Sparkline in Cell I4
// Description: Creates a workbook, inserts sample data in row 4, adds a line sparkline for A4:D4 into cell I4, turns on first‑ and last‑point markers, applies purple and yellow colors, and saves the file as SparklineFirstLastPoint.xlsx using C#.
// Keywords: Aspose.Cells | C# line sparkline | ShowFirstPoint | ShowLastPoint | sparkline point color | Excel sparkline I4 | SparklineGroup | .NET Aspose.Cells | custom sparkline markers | Excel dashboard sparkline
// Common Searches: Aspose.Cells enable first point marker line sparkline | How to show last point in Aspose.Cells sparkline C# | Set custom colors for sparkline markers Aspose.Cells | Add line sparkline to cell I4 using Aspose.Cells | SparklineGroup ShowFirstPoint ShowLastPoint example
// Developer Intent: Insert a line sparkline at I4 and highlight its start and end values with distinct colors.
// Use Cases: Highlight the opening and closing values of a sales trend in a financial report. | Create an Excel dashboard where each row’s sparkline endpoints are colored to draw attention to performance extremes. | Generate automated quality‑control sheets that emphasize the first and last measurements of a process.
// AI Prompts: Generate C# code with Aspose.Cells that adds a line sparkline to I4, shows both first and last markers, and colors them purple and yellow. | Explain how to configure ShowFirstPoint and ShowLastPoint for a SparklineGroup and assign custom CellsColor objects in Aspose.Cells. | Provide an example that creates a sparkline for range A4:D4, places it in cell I4, and uses different colors for the start and end points.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Creates a workbook, inserts sample data in row 4, adds a line sparkline for A4:D4 into cell I4, turns on first‑ and last‑point markers, applies purple and yellow colors, and saves the file as SparklineFirstLastPoint.xlsx using C#.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the sparkline (row 4)
        sheet.Cells["A4"].PutValue(5);
        sheet.Cells["B4"].PutValue(2);
        sheet.Cells["C4"].PutValue(1);
        sheet.Cells["D4"].PutValue(3);

        // Define the location cell I4 (column index 8, row index 3)
        CellArea location = new CellArea
        {
            StartRow = 3,
            EndRow = 3,
            StartColumn = 8,
            EndColumn = 8
        };

        // Add a line sparkline group with the data range A4:D4 and place it at I4
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A4:D4", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Enable highlighting of both the first and last points
        group.ShowFirstPoint = true;
        group.ShowLastPoint = true;

        // Optional: set custom colors for the first and last points
        CellsColor firstPointColor = workbook.CreateCellsColor();
        firstPointColor.Color = Color.Purple;
        group.FirstPointColor = firstPointColor;

        CellsColor lastPointColor = workbook.CreateCellsColor();
        lastPointColor.Color = Color.Yellow;
        group.LastPointColor = lastPointColor;

        // Save the workbook
        workbook.Save("SparklineFirstLastPoint.xlsx");
    }
}
