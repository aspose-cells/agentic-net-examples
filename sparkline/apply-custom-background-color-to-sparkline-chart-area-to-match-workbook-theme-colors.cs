// Title: Set a custom series color for a column sparkline with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set sparkline series color programmatically | how to change column sparkline color using .NET Aspose.Cells | apply theme-matching color to sparkline series in C# workbook | create CellsColor from ARGB and assign to SparklineGroup in Aspose.Cells | example of customizing sparkline series color in an XLSX file with Aspose
// Tags: sparkline series color customization Aspose.Cells | column sparkline group color .NET | create CellsColor from ARGB Aspose | assign series color to SparklineGroup C# | Aspose.Cells sparkline theming example

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineBackgroundDemo
{
    // The sample creates a workbook, fills cells A1:A5 with numbers, adds a column‑type SparklineGroup in B1:B5, builds a CellsColor from an ARGB .NET color, assigns it to the group's SeriesColor, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and obtain the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data for the sparkline.
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["A2"].PutValue(8);
                sheet.Cells["A3"].PutValue(3);
                sheet.Cells["A4"].PutValue(7);
                sheet.Cells["A5"].PutValue(6);

                // Ensure the target cells for sparklines exist.
                for (int row = 0; row < 5; row++)
                {
                    sheet.Cells[row, 1].PutValue(string.Empty);
                }

                // 3. Add a SparklineGroup (type Column) using the data range A1:A5
                //    and placing the sparklines in B1:B5.
                CellArea locationRange = new CellArea
                {
                    StartRow = 0,
                    EndRow = 4,
                    StartColumn = 1,
                    EndColumn = 1
                };

                int sparklineGroupIndex = sheet.SparklineGroups.Add(
                    SparklineType.Column,
                    "A1:A5",
                    false,
                    locationRange);

                SparklineGroup sparkGroup = sheet.SparklineGroups[sparklineGroupIndex];

                // 4. Set a custom series color that matches a theme color.
                //    Use a standard .NET color; Aspose.Cells will handle the conversion.
                CellsColor seriesColor = workbook.CreateCellsColor();
                seriesColor.Color = Color.FromArgb(255, 91, 155, 213); // Example theme-like blue
                sparkGroup.SeriesColor = seriesColor;

                // 5. Save the workbook.
                string outputPath = "SparklineBackgroundDemo.xlsx";

                // Ensure the directory exists before saving.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
