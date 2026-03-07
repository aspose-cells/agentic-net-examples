using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineCopyExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (A1:D1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location of the original sparkline (cell E1)
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

            // Add a sparkline to the original group (placed at E1)
            originalGroup.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // -------------------------------------------------
            // Copy the sparkline group to a new location (cell G1)
            // -------------------------------------------------

            // Define the location for the copied sparkline (cell G1)
            CellArea copyLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 6, // column G
                EndColumn = 6
            };

            // Add a new sparkline group using the same data range and type as the original
            int copyIndex = sheet.SparklineGroups.Add(originalGroup.Type, "A1:D1", false, copyLocation);
            SparklineGroup copyGroup = sheet.SparklineGroups[copyIndex];

            // Copy the sparkline itself (place it at G1)
            copyGroup.Sparklines.Add(sheet.Name + "!A1:D1", 0, 6);

            // Adjust properties of the copied group as required
            // Change the sparkline type to Column
            copyGroup.Type = SparklineType.Column;

            // Set the series color to Red
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Red;
            copyGroup.SeriesColor = seriesColor;

            // Enable markers and set their color to Blue
            copyGroup.ShowMarkers = true;
            CellsColor markersColor = workbook.CreateCellsColor();
            markersColor.Color = Color.Blue;
            copyGroup.MarkersColor = markersColor;

            // Save the workbook to an XLSX file
            workbook.Save("SparklineCopyResult.xlsx", SaveFormat.Xlsx);
        }
    }
}