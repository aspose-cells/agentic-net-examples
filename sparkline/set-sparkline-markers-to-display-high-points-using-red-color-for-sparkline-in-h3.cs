using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineExample
{
    public class SparklineHighPointMarker
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for the sparkline (range A1:D1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed (cell H3)
            CellArea location = new CellArea
            {
                StartColumn = 7, // Column H (0‑based index)
                EndColumn = 7,
                StartRow = 2,    // Row 3 (0‑based index)
                EndRow = 2
            };

            // Add a sparkline group of type Line, using the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group (data range, row index, column index of the location)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 2, 7);

            // Enable high‑point highlighting and set its color to red
            group.ShowHighPoint = true;
            CellsColor highPointColor = workbook.CreateCellsColor();
            highPointColor.Color = Color.Red;
            group.HighPointColor = highPointColor;

            // (Optional) Show markers for each point
            group.ShowMarkers = true;

            // Save the workbook
            workbook.Save("SparklineHighPointMarker.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SparklineHighPointMarker.Run();
        }
    }
}