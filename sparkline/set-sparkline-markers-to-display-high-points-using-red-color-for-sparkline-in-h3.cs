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

        // Sample data for the sparkline (range A1:D1)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(3);

        // Define the location of the sparkline: cell H3 (column index 7, row index 2)
        CellArea sparklineArea = new CellArea
        {
            StartColumn = 7,
            EndColumn = 7,
            StartRow = 2,
            EndRow = 2
        };

        // Add a sparkline group of type Line
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineArea);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the sparkline to the group (optional, but ensures the sparkline exists)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 2, 7);

        // Enable markers and highlight the highest points
        group.ShowMarkers = true;
        group.ShowHighPoint = true;

        // Set the high point color to red
        CellsColor red = workbook.CreateCellsColor();
        red.Color = Color.Red;
        group.HighPointColor = red;

        // Save the workbook
        workbook.Save("SparklineHighPointMarker.xlsx", SaveFormat.Xlsx);
    }
}