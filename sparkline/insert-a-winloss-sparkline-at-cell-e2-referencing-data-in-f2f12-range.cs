// Title: Create a Win/Loss sparkline in cell E2 from range F2:F12 using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to add a win/loss sparkline at E2 referencing F2:F12 | How to generate a win‑loss sparkline in an Excel file with Aspose.Cells .NET | Programmatically insert a win/loss sparkline into a specific cell using Aspose.Cells C# | Aspose.Cells .NET code for creating a win/loss sparkline with custom series color | Add a win‑loss sparkline to an Excel workbook and set its series color with Aspose.Cells
// Tags: sparkline group creation Aspose.Cells .NET | win‑loss sparkline placement C# | custom series color Aspose.Cells sparkline | populate data range for sparkline C# | save workbook with sparkline Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineDemo
{
    // The sample creates a new workbook, fills cells F2:F12 with alternating positive and negative numbers, adds a Win/Loss sparkline group positioned at cell E2, applies a dark blue series color, and saves the file as WinLossSparkline.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Populate sample data in the range F2:F12 (zero‑based rows 1‑11)
                // ------------------------------------------------------------
                for (int i = 0; i < 11; i++)
                {
                    // Alternate positive and negative values to illustrate win/loss
                    double value = (i % 2 == 0) ? (i + 1) : -(i + 1);
                    sheet.Cells[i + 1, 5].PutValue(value); // Column F = index 5
                }

                // ------------------------------------------------------------
                // Define the location cell area for the sparkline (E2)
                // ------------------------------------------------------------
                CellArea location = new CellArea
                {
                    StartRow = 1,   // Row 2 (zero‑based)
                    EndRow = 1,
                    StartColumn = 4, // Column E (zero‑based)
                    EndColumn = 4
                };

                // ------------------------------------------------------------
                // Add a Win/Loss sparkline group
                // ------------------------------------------------------------
                int groupIndex = sheet.SparklineGroups.Add(
                    SparklineType.WinLoss,   // correct enum value
                    "F2:F12",                // data range (A1 style)
                    false,                   // isDateAxis – not a date axis
                    location);               // where the sparkline will be placed

                SparklineGroup group = sheet.SparklineGroups[groupIndex];

                // Add the sparkline to the group (explicitly required)
                // Row and column are zero‑based indices for the target cell (E2)
                int sparklineIdx = group.Sparklines.Add("F2:F12", 1, 4);
                Sparkline sparkline = group.Sparklines[sparklineIdx];

                // Optional: customize appearance (e.g., set series color)
                CellsColor seriesColor = workbook.CreateCellsColor();
                seriesColor.Color = System.Drawing.Color.DarkBlue;
                group.SeriesColor = seriesColor;

                // Ensure the output directory exists
                string outputPath = "WinLossSparkline.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the sparkline:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
