// Title: Copy rows as values only with PasteOptions in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to use Aspose.Cells' PasteOptions with PasteType.Values to copy all populated rows from one worksheet to another while discarding formulas, formatting, and links. The example sets SkipBlanks, configures CopyRows, and saves the destination workbook containing only calculated values.
// Keywords: Aspose.Cells CopyRows values only | PasteOptions PasteType.Values C# | copy rows without formulas Aspose | skip blanks Aspose.Cells | export calculated values Excel .NET | ignore formatting Aspose.Cells copy | PasteOptions example C#
// Common Searches: Aspose.Cells copy rows values only | PasteOptions to ignore formulas in C# | CopyRows method example Aspose.Cells | skip blanks when copying rows Aspose | export only values from worksheet .NET
// Developer Intent: Copy a range of rows from a source worksheet to a destination worksheet, preserving only the evaluated cell values and omitting formulas, formatting, and external links.
// Use Cases: Generate a clean data export for reporting by copying calculation results without any formulas. | Create a summary sheet that contains only final values, eliminating hidden links and style information. | Prepare a dataset for third‑party processing where only raw values are required, with blanks skipped.
// AI Prompts: Show how to copy rows 5‑10 as values only using PasteOptions in Aspose.Cells. | Provide code to copy rows while keeping number formats but removing formulas. | Explain how to retain column widths when copying rows with PasteType.Values.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells' PasteOptions with PasteType.Values to copy all populated rows from one worksheet to another while discarding formulas, formatting, and links. The example sets SkipBlanks, configures CopyRows, and saves the destination workbook containing only calculated values.
class CopyRowsValuesOnly
{
    static void Main()
    {
        // Create source workbook and fill with data, formulas and formatting
        Workbook srcWb = new Workbook();
        Worksheet srcWs = srcWb.Worksheets[0];
        srcWs.Cells["A1"].PutValue(10);
        srcWs.Cells["B1"].Formula = "=A1*2";          // formula that should be ignored
        srcWs.Cells["A2"].PutValue(20);
        srcWs.Cells["B2"].Formula = "=A2*2";
        srcWs.Cells["A3"].PutValue(30);
        srcWs.Cells["B3"].Formula = "=A3*2";

        // Create destination workbook (empty)
        Workbook destWb = new Workbook();
        Worksheet destWs = destWb.Worksheets[0];

        // Configure PasteOptions to copy only values
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Values,   // copy only cell values
            SkipBlanks = true,
            OnlyVisibleCells = false,
            Transpose = false,
            IgnoreLinksToOriginalFile = true
        };

        // Default CopyOptions (no special behavior)
        CopyOptions copyOptions = new CopyOptions();

        // Determine number of rows to copy (all rows that contain data)
        int rowsToCopy = srcWs.Cells.MaxDisplayRange.RowCount;

        // Copy rows from source to destination using the specified options
        destWs.Cells.CopyRows(
            srcWs.Cells,          // source cells
            0,                    // source start row index
            0,                    // destination start row index
            rowsToCopy,           // number of rows to copy
            copyOptions,          // copy options
            pasteOptions);        // paste options (values only)

        // Save the result
        destWb.Save("RowsCopiedValuesOnly.xlsx");
    }
}
