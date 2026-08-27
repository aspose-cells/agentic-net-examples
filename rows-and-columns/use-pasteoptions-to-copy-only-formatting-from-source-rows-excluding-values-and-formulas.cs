// Title: Copy rows with formatting only using PasteOptions (PasteType.Formats) in Aspose.Cells for .NET C#
// AI Prompts: Copy the first three rows from a source worksheet to a destination worksheet while preserving only cell styles using PasteOptions in C#. | Configure PasteOptions with PasteType.Formats to duplicate row formatting without transferring values or formulas in Aspose.Cells. | Generate C# code that copies rows between workbooks, applying only formatting via CopyRows and PasteOptions.
// Common Searches: Aspose.Cells copy rows without values or formulas C# | How to copy only cell formatting between worksheets using Aspose.Cells .NET | PasteOptions PasteType.Formats example for row copy in C# | Copy row styles but not data with Aspose.Cells CopyRows method | Exclude formulas when copying rows in Aspose.Cells for .NET
// Tags: CopyRows with PasteOptions formatting only | PasteOptions PasteType.Formats Aspose.Cells | C# copy row styles without values | Aspose.Cells row formatting transfer | Exclude formulas during row copy .NET

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a source workbook containing a bold header, a yellow‑filled numeric cell, and a formula, then uses CopyRows together with PasteOptions set to PasteType.Formats to copy the first three rows into a new workbook, preserving only the original formatting while omitting values and formulas.
class Program
{
    static void Main()
    {
        // Create source workbook and populate it with data, formulas and formatting
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Header with bold font
        sourceSheet.Cells["A1"].PutValue("Header");
        Style headerStyle = sourceWorkbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        sourceSheet.Cells["A1"].SetStyle(headerStyle);

        // Numeric value with background color
        sourceSheet.Cells["A2"].PutValue(123);
        Style valueStyle = sourceWorkbook.CreateStyle();
        valueStyle.ForegroundColor = Color.Yellow;
        valueStyle.Pattern = BackgroundType.Solid;
        sourceSheet.Cells["A2"].SetStyle(valueStyle);

        // Formula cell
        sourceSheet.Cells["A3"].Formula = "=A2*2";

        // Create destination workbook where rows will be copied
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Default copy options (no special behavior)
        CopyOptions copyOptions = new CopyOptions();

        // Paste options configured to copy only formatting
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Formats   // Formats only, no values or formulas
        };

        // Copy the first three rows (0,1,2) from source to destination starting at row 0
        destinationSheet.Cells.CopyRows(
            sourceSheet.Cells,   // source cells
            0,                   // source row index
            0,                   // destination row index
            3,                   // number of rows to copy
            copyOptions,
            pasteOptions);

        // Save the result
        destinationWorkbook.Save("RowsCopyFormatsOnly.xlsx");
    }
}
