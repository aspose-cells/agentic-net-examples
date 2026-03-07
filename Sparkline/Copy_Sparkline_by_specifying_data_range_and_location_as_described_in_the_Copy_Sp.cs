using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (row 0, columns A‑D)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the original sparkline location (cell E1)
            CellArea originalLocation = CellArea.CreateCellArea("E1", "E1");

            // Add the original sparkline group (Line type) with data range A1:D1
            int originalGroupIdx = sheet.SparklineGroups.Add(
                SparklineType.Line,          // Sparkline type
                "A1:D1",                     // Data range
                false,                       // Plot by row (horizontal)
                originalLocation);           // Location range

            // The group now contains a sparkline at E1
            SparklineGroup originalGroup = sheet.SparklineGroups[originalGroupIdx];

            // -----------------------------------------------------------------
            // Copy the sparkline by creating a new group with the same data range
            // but a different location (cell G1). This demonstrates the "Copy
            // Sparkline by Specifying Data Range and Location" technique.
            // -----------------------------------------------------------------

            // Define the new location for the copied sparkline (cell G1)
            CellArea copyLocation = CellArea.CreateCellArea("G1", "G1");

            // Add a new sparkline group using the same data range as the original
            int copyGroupIdx = sheet.SparklineGroups.Add(
                SparklineType.Line,          // Same sparkline type
                "A1:D1",                     // Same data range
                false,                       // Same orientation
                copyLocation);               // New location range

            // Access the copied group (optional, e.g., to modify appearance)
            SparklineGroup copyGroup = sheet.SparklineGroups[copyGroupIdx];

            // Example: change the series color of the copied sparkline
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = System.Drawing.Color.Blue;
            copyGroup.SeriesColor = seriesColor;

            // Save the workbook to a file
            workbook.Save("SparklineCopyDemo.xlsx");
        }
    }
}