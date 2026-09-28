// Title: Copy a row from one workbook to another while preserving cell formatting using Aspose.Cells PasteOptions in C#
// AI Prompts: Write C# code that copies a specific row from a source worksheet to a destination worksheet and keeps all cell styles by setting PasteOptions.PasteType to All. | Show how to configure CopyOptions and PasteOptions in Aspose.Cells to transfer a styled row between workbooks without losing background colors or fonts. | Provide a complete example that creates a source workbook, applies a yellow background style to a row, copies the row to a new workbook, and saves the file while preserving the style.
// Common Searches: Aspose.Cells C# copy row with formatting between workbooks | how to retain cell background color when copying rows using Aspose.Cells | PasteOptions PasteType.All usage for row copy in .NET | CopyRows method preserving styles Aspose.Cells example | C# Aspose.Cells copy rows without losing style
// Tags: CopyRows with PasteOptions in Aspose.Cells | preserve cell formatting during row copy | PasteType.All for style retention | C# Aspose.Cells row copy between workbooks | apply background color style Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a source workbook, applies a yellow background style to the first row, then uses CopyRows with PasteOptions set to PasteType.All to copy that row into a new workbook while preserving all cell formatting, and finally saves the result as PreserveStyleCopy.xlsx.
class PreserveRowStyleCopy
{
    static void Main()
    {
        // Create source workbook and apply a style to the first row
        Workbook srcWorkbook = new Workbook();
        Worksheet srcSheet = srcWorkbook.Worksheets[0];

        // Fill some data in the first row
        srcSheet.Cells["A1"].PutValue("Styled Text");
        srcSheet.Cells["B1"].PutValue(123);

        // Create a style with a yellow background
        Style rowStyle = srcWorkbook.CreateStyle();
        rowStyle.ForegroundColor = Color.Yellow;
        rowStyle.Pattern = BackgroundType.Solid;

        // Apply the style to the range A1:B1 (first row)
        srcSheet.Cells.CreateRange("A1:B1").SetStyle(rowStyle);

        // Create destination workbook where the row will be copied
        Workbook destWorkbook = new Workbook();
        Worksheet destSheet = destWorkbook.Worksheets[0];

        // Prepare copy and paste options
        CopyOptions copyOptions = new CopyOptions(); // default copy options
        PasteOptions pasteOptions = new PasteOptions
        {
            // Use PasteType.All to preserve all data, including cell styles
            PasteType = PasteType.All
        };

        // Copy the first row from source to destination while preserving styles
        destSheet.Cells.CopyRows(
            srcSheet.Cells,   // source cells
            0,                // source row index (first row)
            0,                // destination row index (first row)
            1,                // number of rows to copy
            copyOptions,      // copy options
            pasteOptions);    // paste options with style preservation

        // Save the resulting workbook
        destWorkbook.Save("PreserveStyleCopy.xlsx");
    }
}
