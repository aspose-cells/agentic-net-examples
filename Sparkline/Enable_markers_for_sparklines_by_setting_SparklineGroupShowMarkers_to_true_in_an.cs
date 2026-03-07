using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineMarkersDemo
{
    class Program
    {
        static void Main(string[] args)
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
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group with the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add a sparkline to the group (the same range, placed at E1)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // Enable markers for each point in the sparkline
            group.ShowMarkers = true;

            // Optional: set the color of the markers
            CellsColor markersColor = workbook.CreateCellsColor();
            markersColor.Color = Color.Black;
            group.MarkersColor = markersColor;

            // Save the workbook to an XLSX file
            workbook.Save("SparklineWithMarkers.xlsx");
        }
    }
}