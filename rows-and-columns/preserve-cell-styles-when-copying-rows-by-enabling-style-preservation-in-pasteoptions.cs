using System;
using System.Drawing;
using Aspose.Cells;

class PreserveRowStyleCopy
{
    static void Main()
    {
        // Create source workbook and add styled data
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Create a style with yellow background
        Style sourceStyle = sourceWorkbook.CreateStyle();
        sourceStyle.ForegroundColor = Color.Yellow;
        sourceStyle.Pattern = BackgroundType.Solid;

        // Apply style to a cell in the first row
        sourceSheet.Cells["A1"].PutValue("Styled Cell");
        sourceSheet.Cells["A1"].SetStyle(sourceStyle);
        sourceSheet.Cells["B1"].PutValue(123); // additional data in the same row

        // Create destination workbook where the row will be copied
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Prepare copy and paste options
        CopyOptions copyOptions = new CopyOptions(); // default options
        PasteOptions pasteOptions = new PasteOptions
        {
            // Preserve all data, formats, and styles
            PasteType = PasteType.All
        };

        // Copy the first row (index 0) from source to destination row index 5
        destinationSheet.Cells.CopyRows(
            sourceSheet.Cells,   // source cells
            0,                   // source row index
            5,                   // destination row index
            1,                   // number of rows to copy
            copyOptions,         // copy options
            pasteOptions);       // paste options with style preservation

        // Save the result workbook
        destinationWorkbook.Save("PreserveStyleCopy.xlsx");
    }
}