// Title: Set chart legend to the right side while keeping its entry fill transparent using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, add a column chart, and move the legend to the right side without altering the default transparent fill of legend entries with Aspose.Cells in C#. | Generate an Excel file containing a column chart where the legend is positioned on the right and the legend entries retain their transparent background, then save the file using the Aspose.Cells Chart and Legend APIs. | Modify an existing chart's legend to the right side while preserving any existing fill transparency settings, employing Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set chart legend position to right side | preserve legend entry fill transparency when changing legend location Aspose.Cells | how to keep legend background transparent in Aspose.Cells chart | move chart legend without affecting fill properties using Aspose.Cells .NET | Aspose.Cells chart legend placement examples C#
// Tags: Aspose.Cells chart legend positioning | C# set legend right side | preserve legend fill transparency | column chart legend configuration .NET | Excel chart legend placement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Demonstrates how to create a workbook, add a column chart, set the legend position to the right side, and keep the legend entries' fill transparency unchanged before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a column chart to the worksheet (from row 5, column 0 to row 20, column 10)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // The Legend class does not have an Angle property; remove or replace with a valid setting if needed
                // Example: set legend position to the right side
                chart.Legend.Position = LegendPositionType.Right;

                // Determine output path
                string outputPath = "output.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

                // Ensure the output directory exists (skip if directory is null or empty)
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
