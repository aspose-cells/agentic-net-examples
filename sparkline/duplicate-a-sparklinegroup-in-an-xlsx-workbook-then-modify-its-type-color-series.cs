using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class DuplicateSparklineGroup
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Assume there is at least one sparkline group in the worksheet
        SparklineGroupCollection groups = sheet.SparklineGroups;
        if (groups.Count == 0)
        {
            Console.WriteLine("No sparkline groups found.");
            return;
        }

        // Get the first (source) sparkline group
        SparklineGroup sourceGroup = groups[0];

        // Retrieve the data range of the first sparkline in the source group
        // (All sparklines in a group share the same data range pattern)
        string dataRange = sourceGroup.Sparklines[0].DataRange;

        // Determine whether the source group is plotted vertically.
        // Aspose.Cells does not expose this directly, so we assume false (by row)
        bool isVertical = false;

        // Define a new location range for the duplicated group.
        // Here we place the duplicated sparklines one column to the right of the original location.
        CellArea sourceLocation = sourceGroup.Sparklines[0].Row >= 0 && sourceGroup.Sparklines[0].Column >= 0
            ? new CellArea
            {
                StartRow = sourceGroup.Sparklines[0].Row,
                EndRow = sourceGroup.Sparklines[0].Row,
                StartColumn = sourceGroup.Sparklines[0].Column + 1,
                EndColumn = sourceGroup.Sparklines[0].Column + 1
            }
            : new CellArea(); // fallback (should not happen)

        // Add a new sparkline group with the same data range and the new location
        int newIndex = groups.Add(sourceGroup.Type, dataRange, isVertical, sourceLocation);
        SparklineGroup newGroup = groups[newIndex];

        // Duplicate each sparkline from the source group into the new group
        foreach (Sparkline srcSparkline in sourceGroup.Sparklines)
        {
            // Add a sparkline to the new group with the same data range.
            // The row is the same as the source; column is shifted by +1 (as defined above).
            newGroup.Sparklines.Add(srcSparkline.DataRange, srcSparkline.Row, sourceLocation.StartColumn);
        }

        // ----- Modify the duplicated group as required -----

        // Change the sparkline type (e.g., to Column)
        newGroup.Type = SparklineType.Column;

        // Set a new series color (e.g., Red)
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Red;
        newGroup.SeriesColor = seriesColor;

        // Enable markers and set their color (e.g., Blue)
        newGroup.ShowMarkers = true;
        CellsColor markersColor = workbook.CreateCellsColor();
        markersColor.Color = Color.Blue;
        newGroup.MarkersColor = markersColor;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}