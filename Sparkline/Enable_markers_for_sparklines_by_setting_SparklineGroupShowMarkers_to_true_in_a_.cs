using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsExamples
{
    public class EnableSparklineMarkers
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that the sparkline will represent
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the cell area where the sparkline will be placed (E1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add a line sparkline group with the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add a sparkline to the group
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // Enable markers for each point in the sparkline
            group.ShowMarkers = true;

            // Set the color of the markers
            CellsColor markersColor = workbook.CreateCellsColor();
            markersColor.Color = Color.Black;
            group.MarkersColor = markersColor;

            // Save the workbook in XLSX format
            workbook.Save("SparklineWithMarkers.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            EnableSparklineMarkers.Run();
        }
    }
}