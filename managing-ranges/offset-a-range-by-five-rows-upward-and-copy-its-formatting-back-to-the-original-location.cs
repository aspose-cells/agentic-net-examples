// Title: Offset a Range Upward by 5 Rows and Copy Its Formatting with Aspose.Cells for .NET
// Description: This C# example demonstrates how to create a workbook, define an original range (A6:D10), generate an offset range five rows above, apply a different style to the offset range, and then use `Range.CopyStyle` to transfer the formatting back to the original cells before saving the file.
// Keywords: Aspose.Cells offset range | copy formatting Aspose.Cells | Range.CopyStyle .NET | C# Aspose.Cells example | shift range rows Aspose | Excel style copy programmatically | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells offset range by rows | How to copy style between ranges in Aspose.Cells | C# move range upward and retain formatting | Range.CopyStyle usage example | Aspose.Cells copy formatting from another range
// Developer Intent: Transfer the formatting of a range that has been moved five rows upward back to its original location using Aspose.Cells for .NET.
// Use Cases: Apply a temporary style to a header area, shift it to data rows for preview, then restore the original cells' appearance with a single call. | Synchronize visual formatting across separate worksheet sections by copying styles from an offset block to the source block. | Implement a “style template” that can be edited in a hidden area and propagated to visible cells without altering their values.
// AI Prompts: Write C# code that offsets a given range by N rows with Aspose.Cells and copies the style back to the original range. | Explain the limitations of `Range.CopyStyle` when copying formatting between non‑contiguous ranges in Aspose.Cells. | Show how to copy only the background color from an offset range to the original range while preserving other style attributes.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // This C# example demonstrates how to create a workbook, define an original range (A6:D10), generate an offset range five rows above, apply a different style to the offset range, and then use `Range.CopyStyle` to transfer the formatting back to the original cells before saving the file.
    public class OffsetRangeCopyFormatting
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // ------------------------------------------------------------
                // Sample data and formatting (for demonstration purposes only)
                // ------------------------------------------------------------
                // Fill the original range (A6:D10) with values and a style
                Style demoStyle = workbook.CreateStyle();
                demoStyle.Font.IsBold = true;
                demoStyle.ForegroundColor = Color.LightYellow;
                demoStyle.Pattern = BackgroundType.Solid;

                // Original range coordinates
                int origStartRow = 5;   // zero‑based index for row 6
                int origStartCol = 0;   // column A
                int rowCount = 5;
                int colCount = 4;       // columns A‑D

                // Populate original range with values and apply the style
                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        cells[origStartRow + r, origStartCol + c].PutValue($"R{r + 1}C{c + 1}");
                        cells[origStartRow + r, origStartCol + c].SetStyle(demoStyle);
                    }
                }

                // ------------------------------------------------------------
                // Define the original range object
                // ------------------------------------------------------------
                AsposeRange originalRange = cells.CreateRange(origStartRow, origStartCol, rowCount, colCount);

                // ------------------------------------------------------------
                // Define the offset range (5 rows upward)
                // ------------------------------------------------------------
                int offsetStartRow = origStartRow - 5; // move up by five rows
                AsposeRange offsetRange = cells.CreateRange(offsetStartRow, origStartCol, rowCount, colCount);

                // (Optional) Change formatting in the offset range to illustrate copying back
                Style offsetStyle = workbook.CreateStyle();
                offsetStyle.Font.IsBold = false;
                offsetStyle.ForegroundColor = Color.LightBlue;
                offsetStyle.Pattern = BackgroundType.Solid;

                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        // Keep the same values, only modify the style
                        cells[offsetStartRow + r, origStartCol + c].SetStyle(offsetStyle);
                    }
                }

                // ------------------------------------------------------------
                // Copy formatting from the offset range back to the original range
                // ------------------------------------------------------------
                originalRange.CopyStyle(offsetRange);

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                string outputPath = "OffsetRangeCopyFormatting.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            OffsetRangeCopyFormatting.Run();
        }
    }
}
