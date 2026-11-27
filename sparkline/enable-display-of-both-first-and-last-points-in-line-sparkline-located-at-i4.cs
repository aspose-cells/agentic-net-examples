// Title: Show first and last points in a line sparkline placed in cell I4 with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# line sparkline show first point marker | C# Aspose.Cells enable last point marker in sparkline | set custom colors for first and last points of a sparkline using Aspose.Cells | add a line sparkline to cell I4 with endpoint markers in Aspose.Cells .NET | Aspose.Cells SparklineGroup ShowFirstPoint and ShowLastPoint example
// Tags: line sparkline first point visibility Aspose.Cells | sparklinegroup showlastpoint property C# | custom endpoint colors sparkline Aspose.Cells | sparkline placement cell I4 Aspose.Cells | enable endpoint markers sparkline .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace SparklineFirstLastPointDemo
{
    // The example creates a workbook, fills row 4 with numeric data, adds a line sparkline in cell I4 based on range A4:H4, enables the first and last point markers, assigns green to the first point and red to the last point, and saves the file as SparklineFirstLastPoint.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in row 4 (index 3) from column A to H
            // This data will be used by the sparkline
            for (int col = 0; col < 8; col++)
            {
                sheet.Cells[3, col].PutValue((col + 1) * 10); // Example values: 10,20,...,80
            }

            // Define the cell where the sparkline will be placed (I4)
            // I column = index 8, row 4 = index 3
            CellArea sparklineLocation = CellArea.CreateCellArea("I4", "I4");

            // Add a line sparkline group that uses the data range A4:H4
            // The third parameter (false) indicates that the data is plotted left‑to‑right
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,          // Line sparkline
                "A4:H4",                     // Data range
                false,                       // Plot right‑to‑left = false
                sparklineLocation);          // Where the sparkline will appear

            // Retrieve the created sparkline group
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Enable display of the first and last points
            group.ShowFirstPoint = true;
            group.ShowLastPoint = true;

            // (Optional) Set colors for the first and last points to make them visible
            group.FirstPointColor = workbook.CreateCellsColor();
            group.FirstPointColor.Color = Color.Green;   // First point in green

            group.LastPointColor = workbook.CreateCellsColor();
            group.LastPointColor.Color = Color.Red;      // Last point in red

            // Save the workbook
            string outputPath = "SparklineFirstLastPoint.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'. Sparkline at I4 now shows first and last points.");
        }
    }
}
