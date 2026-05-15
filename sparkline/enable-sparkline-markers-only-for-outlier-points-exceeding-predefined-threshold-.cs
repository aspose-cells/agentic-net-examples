using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineOutlierMarkersDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data – some values exceed the outlier threshold (e.g., > 8)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(12); // outlier
            sheet.Cells["D1"].PutValue(3);
            sheet.Cells["E1"].PutValue(15); // outlier
            sheet.Cells["F1"].PutValue(4);

            // Define the location where the sparkline will be placed (column G)
            CellArea sparklineLocation = new CellArea
            {
                StartColumn = 6, // column G (0‑based index)
                EndColumn = 6,
                StartRow = 0,
                EndRow = 0
            };

            // Add a sparkline group for the data range A1:F1
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:F1", false, sparklineLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add the sparkline to the group (the Add method also creates the sparkline)
            group.Sparklines.Add(sheet.Name + "!A1:F1", 0, 6);

            // Enable markers – this will show a marker for every point
            group.ShowMarkers = true;

            // Set a distinct color for the markers (e.g., red) to make them stand out
            CellsColor markersColor = workbook.CreateCellsColor();
            markersColor.Color = Color.Red;
            group.MarkersColor = markersColor;

            // Aspose.Cells does not provide a direct way to show markers only for
            // points that exceed a custom threshold. The ShowMarkers property applies
            // to all points in the sparkline. As a workaround, you can highlight the
            // highest/lowest points using ShowHighPoint / ShowLowPoint, but these are
            // based on the data extremes, not an arbitrary threshold.
            // Example of highlighting the highest points (which may coincide with outliers):
            group.ShowHighPoint = true;
            CellsColor highPointColor = workbook.CreateCellsColor();
            highPointColor.Color = Color.Orange;
            group.HighPointColor = highPointColor;

            // Save the workbook
            workbook.Save("SparklineOutlierMarkersDemo.xlsx");
        }
    }
}