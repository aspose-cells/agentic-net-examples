// Title: Reposition an Excel slicer to cell D5 while preserving its size using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# code to place a slicer at cell D5 | how to set slicer location based on cell address using Aspose.Cells | preserve slicer width and height while moving it in a .NET workbook | align Excel slicer with existing charts programmatically with Aspose.Cells | convert Excel cell coordinates to pixels for object positioning in C#
// Tags: set slicer location using pixel offsets Aspose.Cells | retain slicer size during repositioning C# | cell-to-pixel conversion for drawing objects Aspose.Cells | align slicer with chart objects programmatically | worksheet slicer manipulation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Example
{
    // The example loads a workbook, finds the first slicer, calculates the pixel coordinates of cell D5, moves the slicer to that position while keeping its original width and height, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Work with the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Locate the first slicer (use dynamic to avoid compile‑time dependency on the Slicer type)
                dynamic slicer = null;
                if (worksheet.Slicers.Count > 0)
                {
                    slicer = worksheet.Slicers[0];
                }
                else
                {
                    Console.WriteLine("No slicer found in the worksheet.");
                    // Continue without slicer manipulation
                }

                if (slicer != null)
                {
                    // Target cell for the slicer position
                    Cell targetCell = worksheet.Cells["D5"];
                    int targetRow = targetCell.Row;
                    int targetColumn = targetCell.Column;

                    // Convert cell's top‑left corner to pixel coordinates
                    int leftPixel = worksheet.Cells.GetColumnWidthPixel(targetColumn);
                    int topPixel = worksheet.Cells.GetRowHeightPixel(targetRow);

                    // Preserve current size while repositioning
                    slicer.SetPosition(leftPixel, topPixel, slicer.Width, slicer.Height);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
