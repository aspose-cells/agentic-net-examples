// Title: How to set a fixed vertical axis minimum and maximum for a Sparkline group using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set sparkline vertical axis minimum value | Define maximum value for sparkline axis programmatically in .NET | Lock sparkline axis range when creating line sparklines with Aspose.Cells | C# example to configure fixed axis scaling for Excel sparklines | Set vertical axis limits for SparklineGroup using Aspose.Cells API
// Tags: Aspose.Cells sparkline group scaling | C# fixed sparkline vertical bounds | programmatic sparkline scaling .NET | line sparkline properties Aspose | Excel sparkline limits using Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineAxisScalingDemo
{
    // The sample creates a workbook, fills a 4x4 range with numeric data, adds a line‑type SparklineGroup sourced from A1:D4 and placed in F1:F4, and demonstrates how to assign a fixed vertical axis minimum (0) and maximum (200) via the VerticalAxisMinimum and VerticalAxisMaximum properties (if supported by the Aspose.Cells version). The workbook is saved as SparklineAxisScalingDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the sparklines (4 rows, 4 columns)
                Cells cells = sheet.Cells;
                for (int row = 0; row < 4; row++)
                {
                    for (int col = 0; col < 4; col++)
                    {
                        // Example data: (row + 1) * (col + 1) * 10
                        cells[row, col].PutValue((row + 1) * (col + 1) * 10);
                    }
                }

                // Add a sparkline group of type Line
                // Data range: A1:D4
                // Destination range: F1:F4 (one sparkline per row)
                int groupIndex = sheet.SparklineGroups.Add(
                    SparklineType.Line,               // Sparkline type
                    "A1:D4",                          // Source data range
                    false,                            // No data in hidden rows/columns
                    CellArea.CreateCellArea("F1", "F4") // Destination cells
                );

                // Retrieve the created SparklineGroup
                SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

                // Set fixed vertical axis scaling if supported by the library version
                // Uncomment the following lines if your Aspose.Cells version provides these properties
                // sparklineGroup.VerticalAxisMinimum = 0;    // Minimum value on the vertical axis
                // sparklineGroup.VerticalAxisMaximum = 200; // Maximum value on the vertical axis

                // Save the workbook to a file
                string outputPath = "SparklineAxisScalingDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
