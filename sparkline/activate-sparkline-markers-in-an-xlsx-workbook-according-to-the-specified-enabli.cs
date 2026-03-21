using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class ActivateSparklineMarkers
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell area where the sparkline will be placed (E1)
        CellArea location = new CellArea
        {
            StartColumn = 4, // Column E (0‑based index)
            EndColumn = 4,
            StartRow = 0,
            EndRow = 0
        };

        // Add a line sparkline group with the data range A1:D1
        int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIdx];

        // Enable markers for each point in the sparkline
        group.ShowMarkers = true;

        // Optionally set the color of the markers
        CellsColor markerColor = workbook.CreateCellsColor();
        markerColor.Color = Color.Black;
        group.MarkersColor = markerColor;

        // Save the workbook with the configured sparkline
        workbook.Save("SparklineWithMarkers.xlsx");
    }
}