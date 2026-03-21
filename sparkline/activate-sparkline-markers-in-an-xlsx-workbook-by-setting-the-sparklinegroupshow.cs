using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class ActivateSparklineMarkers
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the location where the sparkline will be placed (cell E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group of type Line with the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group (the Add method already creates one, this is optional)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Activate markers for each point in the sparkline
        group.ShowMarkers = true;

        // Optionally set the color of the markers
        CellsColor markersColor = workbook.CreateCellsColor();
        markersColor.Color = Color.Black;
        group.MarkersColor = markersColor;

        // Save the workbook with the sparkline markers enabled
        workbook.Save("SparklineWithMarkers.xlsx");
    }
}