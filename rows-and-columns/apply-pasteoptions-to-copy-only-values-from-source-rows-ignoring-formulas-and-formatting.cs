// Title: Copy Row Values Only with PasteOptions in Aspose.Cells for C#/.NET
// Description: Shows how to copy the first two rows from columns A‑B to columns C‑D using Aspose.Cells PasteOptions that transfer just the calculated data, omitting formulas, styles, and empty cells, then saves the workbook.
// Keywords: Aspose.Cells | PasteOptions | values-only copy | ignore formulas | skip blanks | C# sample | .NET Excel | range duplication
// Common Searches: Aspose.Cells copy values C# | PasteOptions values only example | Ignore formulas when copying rows Aspose | Copy range without formatting .NET | Skip blank cells PasteOptions
// Developer Intent: Move rows between worksheets while keeping only the numeric results and discarding any formulas or formatting.
// Use Cases: Export calculation results for clients without revealing formulas | Create a summary sheet that contains only final numbers | Replicate data across multiple sheets while preserving each sheet's own styling
// AI Prompts: Generate C# code that uses Aspose.Cells to copy a range of rows to another area, preserving only the cell values and ignoring blanks. | Describe the PasteOptions settings required in Aspose.Cells to copy values only and exclude formulas and styles.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Shows how to copy the first two rows from columns A‑B to columns C‑D using Aspose.Cells PasteOptions that transfer just the calculated data, omitting formulas, styles, and empty cells, then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a source workbook and fill it with data, formulas and formatting
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Row 1
            sourceSheet.Cells["A1"].PutValue(10);                 // plain value
            sourceSheet.Cells["B1"].Formula = "=A1*2";            // formula
            // Row 2
            sourceSheet.Cells["A2"].PutValue(20);
            sourceSheet.Cells["B2"].Formula = "=A2*2";

            // Apply a simple format to the formula cells (red font)
            Style redStyle = sourceWorkbook.CreateStyle();
            redStyle.Font.Color = Color.Red;
            sourceSheet.Cells["B1"].SetStyle(redStyle);
            sourceSheet.Cells["B2"].SetStyle(redStyle);

            // Create a destination workbook where the values will be copied
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Define the source range (first two rows, columns A and B)
            AsposeRange sourceRange = sourceSheet.Cells.CreateRange(0, 0, 2, 2);
            // Define the destination range (same size, starting at column C)
            AsposeRange destinationRange = destinationSheet.Cells.CreateRange(0, 2, 2, 2);

            // Configure PasteOptions to copy only the values (ignore formulas and formatting)
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.Values,          // copy values only
                SkipBlanks = true,                     // skip blank cells
                OnlyVisibleCells = false,              // include hidden cells
                Transpose = false,                     // no transposition
                IgnoreLinksToOriginalFile = true       // ignore external links
            };

            // Perform the copy using the defined paste options
            destinationRange.Copy(sourceRange, pasteOptions);

            // Save the result
            destinationWorkbook.Save("CopyValuesOnly.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
