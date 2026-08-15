// Title: Copy a range with data and formatting to another worksheet using PasteOptions (Aspose.Cells for .NET)
// Description: Demonstrates how to create a workbook, fill a 5×5 range, apply a bold yellow style, and copy the entire range—including values, formulas, and cell formatting—to a different worksheet with PasteOptions.PasteType = All, then save the file.
// Keywords: Aspose.Cells copy range | PasteOptions All | retain cell formatting | C# range copy worksheet | Aspose.Cells .NET example | duplicate styled range
// Common Searches: Aspose.Cells copy range with formatting | PasteOptions retain styles C# | copy cells between worksheets .NET | how to duplicate a styled range Aspose.Cells | copy range values and formats Aspose
// Developer Intent: Transfer a source range to a destination range on another worksheet while preserving values, formulas, and all visual styles.
// Use Cases: Clone a formatted header block to new report sheets. | Create template sheets and replicate them for each department. | Archive a styled data section in a separate worksheet for backup.
// AI Prompts: Show C# code that copies a range with all formatting using Aspose.Cells PasteOptions. | Explain how PasteOptions.PasteType.All preserves styles when moving cells between worksheets. | Provide an example of duplicating a styled 5x5 block to another sheet in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    // Demonstrates how to create a workbook, fill a 5×5 range, apply a bold yellow style, and copy the entire range—including values, formulas, and cell formatting—to a different worksheet with PasteOptions.PasteType = All, then save the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet (source)
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Populate source range with data
                Cells srcCells = sourceSheet.Cells;
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        srcCells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Create a style (bold font with yellow background) and apply to the source range
                Style srcStyle = workbook.CreateStyle();
                srcStyle.Font.IsBold = true;
                srcStyle.ForegroundColor = System.Drawing.Color.Yellow;
                srcStyle.Pattern = BackgroundType.Solid;
                AsposeRange srcRange = srcCells.CreateRange(0, 0, 5, 5);
                srcRange.SetStyle(srcStyle);

                // Add a second worksheet (destination)
                Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                destSheet.Name = "Destination";

                // Define destination range (same size as source)
                Cells destCells = destSheet.Cells;
                AsposeRange destRange = destCells.CreateRange(0, 0, 5, 5);

                // Configure paste options to copy everything (values, formulas, formats, etc.)
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All // ensures data and formatting are retained
                };

                // Perform the copy from source range to destination range using the paste options
                destRange.Copy(srcRange, pasteOptions);

                // Save the workbook
                string outputPath = "RangeCopyWithStyles.xlsx";
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
