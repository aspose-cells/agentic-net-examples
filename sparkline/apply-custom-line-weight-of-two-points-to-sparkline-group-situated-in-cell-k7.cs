using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsSparklineLineWeight
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for the sparkline (adjust as needed)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location range for the sparkline group (cell K7)
            // Column K -> index 10, Row 7 -> index 6 (zero‑based)
            CellArea location = new CellArea
            {
                StartColumn = 10,
                EndColumn = 10,
                StartRow = 6,
                EndRow = 6
            };

            // Add a line sparkline group with the data range A1:D1 and place it at K7
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group (required when using Add with location)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 10);

            // Set custom line weight to 2 points
            group.LineWeight = 2.0;

            // Optional: customize appearance (e.g., series color) for better visibility
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            group.SeriesColor = seriesColor;

            // Save the workbook
            workbook.Save("SparklineLineWeightK7.xlsx", SaveFormat.Xlsx);
        }
    }
}