// Title: Copy a Range with Data and Formatting to Another Worksheet using PasteOptions (Aspose.Cells for .NET)
// Description: Demonstrates how to create a workbook, fill a 5 × 5 range with values and a custom style, and then copy that range to a different worksheet while preserving values, formulas, and all formatting by using PasteOptions.PasteType.All. The result is saved as an XLSX file.
// Keywords: Aspose.Cells copy range | PasteOptions preserve formatting | C# copy cells with styles | copy range to another worksheet | Aspose.Cells .NET example | retain cell formatting Aspose
// Common Searches: Aspose.Cells copy range with formatting | How to retain styles when copying cells in C# | PasteOptions All copy data and formatting | Copy range between worksheets Aspose.Cells | C# example for copying styled range
// Developer Intent: Copy a source range to a destination range on a different worksheet while keeping all values, formulas, and visual styles intact.
// Use Cases: Duplicate a formatted report template from a master sheet to new period sheets. | Move a styled data block to a summary sheet without losing colors, fonts, or number formats. | Clone a calculation area with its formulas and appearance for testing or backup.
// AI Prompts: Generate C# code that copies a range using Aspose.Cells but only retains values and number formats. | Show how to copy a range that includes conditional formatting with PasteOptions in Aspose.Cells. | Explain the steps to copy a range to a different workbook while preserving all styles using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, fill a 5 × 5 range with values and a custom style, and then copy that range to a different worksheet while preserving values, formulas, and all formatting by using PasteOptions.PasteType.All. The result is saved as an XLSX file.
    public class CopyRangeWithOptionsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Source worksheet
                Worksheet srcSheet = workbook.Worksheets[0];
                srcSheet.Name = "Source";

                // Populate source range with data and a simple style
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        Cell cell = srcSheet.Cells[row, col];
                        cell.PutValue($"R{row}C{col}");

                        // Apply a style to demonstrate formatting copy
                        Style style = workbook.CreateStyle();
                        style.Font.Name = "Arial";
                        style.Font.Size = 12;
                        style.Font.IsBold = (row + col) % 2 == 0;
                        style.ForegroundColor = Color.LightYellow;
                        style.Pattern = BackgroundType.Solid;
                        cell.SetStyle(style);
                    }
                }

                // Destination worksheet
                Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                destSheet.Name = "Destination";

                // Define source and destination ranges (both 5x5 starting at A1)
                AsposeRange srcRange = srcSheet.Cells.CreateRange(0, 0, 5, 5);
                AsposeRange destRange = destSheet.Cells.CreateRange(0, 0, 5, 5);

                // Configure paste options to retain all data and formatting
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All // copies values, formulas, formats, etc.
                };

                // Copy the source range to the destination range using the options
                destRange.Copy(srcRange, pasteOptions);

                // Save the workbook
                string outputPath = "CopyRangeWithStyles.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            CopyRangeWithOptionsDemo.Run();
        }
    }
}
