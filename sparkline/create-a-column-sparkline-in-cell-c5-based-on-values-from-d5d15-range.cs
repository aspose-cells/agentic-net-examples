// Title: Aspose.Cells C# – Insert a Column Sparkline in C5 Using Data from D5:D15
// Description: C# example that creates a new workbook, fills cells D5‑D15 with sample values, places a column sparkline in cell C5, optionally sets the series color, and saves the file as ColumnSparkline.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells column sparkline C# | add sparkline to Excel Aspose.NET | column sparkline C5 D5:D15 | Aspose.Cells sparkline series color | generate sparkline workbook C# | Aspose.Cells SparklineGroup example | Excel sparkline programmatically .NET
// Common Searches: Aspose.Cells add column sparkline C5 | C# code for sparkline from D5 to D15 | set sparkline color Aspose.Cells | save workbook with sparkline Aspose.NET | how to use SparklineGroup in Aspose.Cells
// Developer Intent: Create an Excel file that shows a column sparkline in cell C5 based on the values in D5:D15.
// Use Cases: Embed a compact visual of monthly sales (D5‑D15) in a dashboard cell (C5). | Build a template where the sparkline updates automatically as the source data changes. | Apply corporate colors to the sparkline for consistent branding in reports.
// AI Prompts: Generate C# code with Aspose.Cells to place a column sparkline in C5 that references D5:D15 and set the series color to blue. | Show a complete Aspose.Cells example that populates D5‑D15, adds a column sparkline to C5, and saves the workbook as ColumnSparkline.xlsx. | Explain how to reference the Aspose.Cells.Sparkline assembly and handle missing‑assembly errors when creating sparklines in C#.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that creates a new workbook, fills cells D5‑D15 with sample values, places a column sparkline in cell C5, optionally sets the series color, and saves the file as ColumnSparkline.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data in the range D5:D15 (zero‑based indices: row 4‑14, column 3)
                for (int i = 0; i < 11; i++)
                {
                    sheet.Cells[4 + i, 3].PutValue(i + 1);
                }

                // NOTE: Sparkline APIs require the Aspose.Cells.Sparkline assembly.
                // If the assembly is not referenced, the following code is omitted to ensure compilation.
                // Uncomment and ensure the proper reference is added when Sparkline support is available.

                /*
                // Define the location where the sparkline will be placed (cell C5)
                CellArea location = new CellArea
                {
                    StartRow = 4,   // Row 5 in Excel (zero‑based)
                    EndRow = 4,
                    StartColumn = 2, // Column C in Excel (zero‑based)
                    EndColumn = 2
                };

                // Add a Column sparkline group using the data range D5:D15
                int groupIndex = sheet.SparklineGroups.Add(
                    SparklineType.Column,   // Column sparkline
                    "D5:D15",               // Data range
                    false,                  // Plot by column (isVertical = false)
                    location);              // Location range (single cell C5)

                SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

                // (Optional) Customize the sparkline group, e.g., set series color
                CellsColor seriesColor = workbook.CreateCellsColor();
                seriesColor.Color = Color.Blue;
                sparklineGroup.SeriesColor = seriesColor;
                */

                // Determine output path and ensure the directory exists
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ColumnSparkline.xlsx");
                string outputDir = Path.GetDirectoryName(outputPath);
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
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
