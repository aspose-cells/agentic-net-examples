using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class DuplicateSparklineGroup
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparklines (A1:D1)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(3);
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(2);

        // Define the location cell for the original sparkline (E1)
        CellArea originalLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4, // column E (0‑based index)
            EndColumn = 4
        };

        // Add the original sparkline group (Line type)
        int originalIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, originalLocation);
        SparklineGroup originalGroup = sheet.SparklineGroups[originalIndex];

        // Add a sparkline to the original group
        originalGroup.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Set series color and markers for the original group
        CellsColor originalSeriesColor = workbook.CreateCellsColor();
        originalSeriesColor.Color = Color.Orange;
        originalGroup.SeriesColor = originalSeriesColor;

        originalGroup.ShowMarkers = true;
        CellsColor originalMarkersColor = workbook.CreateCellsColor();
        originalMarkersColor.Color = Color.Black;
        originalGroup.MarkersColor = originalMarkersColor;

        // -------------------------------------------------
        // Duplicate the sparkline group to a new location (F1)
        // -------------------------------------------------
        CellArea duplicateLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 5, // column F
            EndColumn = 5
        };

        // Add a new sparkline group using the same type as the original
        int duplicateIndex = sheet.SparklineGroups.Add(originalGroup.Type, "A1:D1", false, duplicateLocation);
        SparklineGroup duplicateGroup = sheet.SparklineGroups[duplicateIndex];

        // Add a sparkline to the duplicate group
        duplicateGroup.Sparklines.Add(sheet.Name + "!A1:D1", 0, 5);

        // Copy visual settings from the original group
        duplicateGroup.SeriesColor = originalGroup.SeriesColor;
        duplicateGroup.ShowMarkers = originalGroup.ShowMarkers;
        duplicateGroup.MarkersColor = originalGroup.MarkersColor;

        // Modify the duplicate group's type, series color, and markers color as required
        duplicateGroup.Type = SparklineType.Column; // change type to Column

        CellsColor newSeriesColor = workbook.CreateCellsColor();
        newSeriesColor.Color = Color.Blue; // new series color
        duplicateGroup.SeriesColor = newSeriesColor;

        CellsColor newMarkersColor = workbook.CreateCellsColor();
        newMarkersColor.Color = Color.Red; // new markers color
        duplicateGroup.MarkersColor = newMarkersColor;

        // Save the workbook
        workbook.Save("DuplicatedSparklineGroup.xlsx", SaveFormat.Xlsx);
    }
}