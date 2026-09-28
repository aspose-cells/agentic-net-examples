// Title: Apply a custom theme color as the background of a sparkline cell with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a line sparkline, sets a custom theme color for the series, and fills the sparkline’s containing cell with the same solid color using Aspose.Cells. | Show how to use Aspose.Cells to style a SparklineGroup and apply a solid background fill to the target cell that matches a workbook theme color. | Write a .NET example that adds a sparkline to a worksheet, defines a custom CellsColor, assigns it to SeriesColor, and applies it as the cell’s ForegroundColor with a solid pattern.
// Common Searches: aspnet set sparkline cell fill color using Aspose.Cells | c# change sparkline background to match workbook theme | how to apply custom series color and cell background for sparkline in Excel with Aspose | Aspose.Cells example for solid fill of sparkline location | line sparkline with theme‑based background color in C#
// Tags: sparkline cell background fill Aspose.Cells | custom series color line sparkline .NET | apply theme color to SparklineGroup | solid pattern style for sparkline cell | Aspose.Cells workbook theme color usage

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineBackgroundDemo
{
    // The example creates a workbook, writes sample data to A1‑A5, adds a line sparkline in B1, defines a custom blue CellsColor, assigns it to the sparkline's SeriesColor, and applies the same color as a solid fill to the cell's style before saving the file as SparklineBackgroundDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the sparkline
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["A2"].PutValue(3);
                sheet.Cells["A3"].PutValue(8);
                sheet.Cells["A4"].PutValue(2);
                sheet.Cells["A5"].PutValue(7);

                // Define the location where the sparkline will be displayed (column B, row 1)
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 1,
                    EndColumn = 1
                };

                // Add a sparkline group of type Line
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:A5", false, location);
                SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

                // Set the series color to a custom theme‑like blue shade
                CellsColor seriesColor = workbook.CreateCellsColor();
                seriesColor.Color = Color.FromArgb(79, 129, 189);
                sparklineGroup.SeriesColor = seriesColor;

                // Optionally, set the background color of the cell containing the sparkline
                Style style = workbook.CreateStyle();
                style.ForegroundColor = seriesColor.Color; // ForegroundColor expects System.Drawing.Color
                style.Pattern = BackgroundType.Solid;
                sheet.Cells["B1"].SetStyle(style);

                // Save the workbook
                workbook.Save("SparklineBackgroundDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
