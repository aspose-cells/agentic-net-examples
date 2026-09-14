// Title: Show first and last point markers in a line sparkline placed in cell I4 using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a line sparkline from A1:D1, positions it in cell I4, and enables first and last point markers. | Set the SparklineGroup to color the first point purple and the last point yellow. | Save the workbook as SparklineFirstLastPoint.xlsx after configuring the sparkline settings.
// Common Searches: Aspose.Cells C# how to display first and last markers in a line sparkline at a specific cell | set custom colors for sparkline first and last points using Aspose.Cells .NET | add a line sparkline to cell I4 based on range A1:D1 with Aspose.Cells
// Tags: Aspose.Cells sparkline first point marker | Aspose.Cells sparkline last point marker | SparklineGroup ShowFirstPoint property | SparklineGroup ShowLastPoint property | Aspose.Cells sparkline point color customization | place sparkline in specific cell Aspose.Cells

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, inserts sample data, adds a line sparkline at cell I4 sourced from A1:D1, enables first and last point markers, applies purple and yellow colors to those markers, and saves the file as SparklineFirstLastPoint.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the sparkline (row 1, columns A‑D)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the location of the sparkline: cell I4 (zero‑based column 8, row 3)
        CellArea location = new CellArea
        {
            StartRow = 3,
            EndRow = 3,
            StartColumn = 8,
            EndColumn = 8
        };

        // Add a line sparkline group with the data range A1:D1 and place it at I4
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Enable highlighting of the first and last points
        group.ShowFirstPoint = true;
        group.ShowLastPoint = true;

        // Optional: set colors for the first and last points
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
