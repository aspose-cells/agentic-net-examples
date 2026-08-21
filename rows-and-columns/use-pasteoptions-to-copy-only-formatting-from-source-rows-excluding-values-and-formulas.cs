// Title: Copy Row Formatting Only with PasteOptions in Aspose.Cells (C#/.NET)
// Description: C# example that copies only the formatting of rows from a source worksheet to a destination workbook using Aspose.Cells. It sets PasteOptions.PasteType to PasteType.Formats and calls Cells.CopyRows, so styles are transferred while values and formulas are excluded. The sample creates a styled source file, copies the rows, and saves both workbooks.
// Keywords: Aspose.Cells | C# | .NET | CopyRows | PasteOptions | PasteType.Formats | row formatting | copy formatting only | exclude values | exclude formulas | Excel workbook styling | cell style transfer | Aspose.Cells example | copy row styles | Excel automation
// Common Searches: Aspose.Cells copy row formatting only | PasteOptions format only C# | CopyRows without values Aspose | How to copy only styles between worksheets Aspose.Cells | Transfer row styles to another workbook .NET | Copy formatting of rows Aspose.Cells | PasteType.Formats example C#
// Developer Intent: Transfer row styles from one worksheet to another while leaving cell values and formulas untouched.
// Use Cases: Apply a pre‑designed header style to multiple data workbooks without copying the data. | Generate reports that reuse existing row formatting while programmatically inserting fresh content. | Synchronize the visual layout of separate Excel files when only appearance needs to be shared.
// AI Prompts: Show a C# snippet that copies only row formatting between two Aspose.Cells worksheets using PasteOptions. | Generate an Aspose.Cells example that copies a specific range of rows with formatting only, preserving no cell values. | Explain how to modify the code to copy column formatting instead of rows while using PasteType.Formats.

using System;
using Aspose.Cells;
using System.Drawing;

// C# example that copies only the formatting of rows from a source worksheet to a destination workbook using Aspose.Cells. It sets PasteOptions.PasteType to PasteType.Formats and calls Cells.CopyRows, so styles are transferred while values and formulas are excluded. The sample creates a styled source file, copies the rows, and saves both workbooks.
class CopyRowFormattingOnly
{
    static void Main()
    {
        // Create source workbook and add data with formatting
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Row 0: set value and make font bold
        sourceSheet.Cells["A1"].PutValue("Header");
        Style headerStyle = sourceWorkbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = Color.White;
        headerStyle.ForegroundColor = Color.DarkBlue;
        headerStyle.Pattern = BackgroundType.Solid;
        sourceSheet.Cells["A1"].SetStyle(headerStyle);

        // Row 1: set value and apply background color
        sourceSheet.Cells["A2"].PutValue("Data 1");
        Style dataStyle = sourceWorkbook.CreateStyle();
        dataStyle.ForegroundColor = Color.LightYellow;
        dataStyle.Pattern = BackgroundType.Solid;
        sourceSheet.Cells["A2"].SetStyle(dataStyle);

        // Row 2: set value and apply another background color
        sourceSheet.Cells["A3"].PutValue("Data 2");
        Style dataStyle2 = sourceWorkbook.CreateStyle();
        dataStyle2.ForegroundColor = Color.LightGreen;
        dataStyle2.Pattern = BackgroundType.Solid;
        sourceSheet.Cells["A3"].SetStyle(dataStyle2);

        // Create destination workbook (empty)
        Workbook destWorkbook = new Workbook();
        Worksheet destSheet = destWorkbook.Worksheets[0];

        // Prepare copy and paste options
        CopyOptions copyOptions = new CopyOptions(); // default options
        PasteOptions pasteOptions = new PasteOptions
        {
            // Copy only formatting, no values or formulas
            PasteType = PasteType.Formats
        };

        // Copy the rows from source to destination using the options
        // Here we copy all rows that contain data in the source sheet
        int rowsToCopy = sourceSheet.Cells.MaxDisplayRange.RowCount;
        destSheet.Cells.CopyRows(
            sourceSheet.Cells,          // source cells
            0,                          // source start row index
            0,                          // destination start row index
            rowsToCopy,                 // number of rows to copy
            copyOptions,                // copy options (default)
            pasteOptions);              // paste options (formats only)

        // Save both workbooks for verification
        sourceWorkbook.Save("SourceWorkbook.xlsx");
        destWorkbook.Save("DestinationWorkbook_FormattingOnly.xlsx");
    }
}
