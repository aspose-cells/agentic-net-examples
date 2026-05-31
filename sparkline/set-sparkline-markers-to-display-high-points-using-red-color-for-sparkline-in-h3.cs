using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineHighPointMarker
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for the sparkline (adjust range as needed)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(3);

        // Define the location of the sparkline (cell H3)
        CellArea sparklineArea = new CellArea
        {
            StartColumn = 7, // Column H (0‑based index)
            EndColumn = 7,
            StartRow = 2,    // Row 3 (0‑based index)
            EndRow = 2
        };

        // Add a line sparkline group that uses the data range A1:D1 and places the sparkline in H3
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineArea);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the sparkline to the group (required step)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 2, 7);

        // Enable high‑point markers and set their color to red
        group.ShowHighPoint = true;
        CellsColor redColor = workbook.CreateCellsColor();
        redColor.Color = Color.Red;
        group.HighPointColor = redColor;

        // (Optional) Show markers for all points; they will inherit the high‑point color for high values
        group.ShowMarkers = true;
        group.MarkersColor = redColor;

        // Save the workbook
        workbook.Save("SparklineHighPointMarker.xlsx", SaveFormat.Xlsx);
    }
}