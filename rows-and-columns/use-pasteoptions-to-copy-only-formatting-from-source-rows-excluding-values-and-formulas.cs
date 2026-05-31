using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create source workbook and add some data with formatting
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Header row with bold font
        sourceSheet.Cells["A1"].PutValue("Header");
        Style headerStyle = sourceWorkbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        sourceSheet.Cells["A1"].SetStyle(headerStyle);

        // Data row with background color
        sourceSheet.Cells["A2"].PutValue(123);
        Style dataStyle = sourceWorkbook.CreateStyle();
        dataStyle.ForegroundColor = Color.Yellow;
        dataStyle.Pattern = BackgroundType.Solid;
        sourceSheet.Cells["A2"].SetStyle(dataStyle);

        // Create destination workbook
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // CopyOptions (default settings)
        CopyOptions copyOptions = new CopyOptions();

        // PasteOptions configured to copy only formats
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Formats
        };

        // Copy the rows from source to destination using the overload that accepts both CopyOptions and PasteOptions
        // Here we copy all rows that contain data in the source sheet
        int rowsToCopy = sourceSheet.Cells.MaxDisplayRange.RowCount;
        destinationSheet.Cells.CopyRows(
            sourceSheet.Cells,
            0,                 // source start row index
            0,                 // destination start row index
            rowsToCopy,        // number of rows to copy
            copyOptions,
            pasteOptions);

        // Save the result workbook
        destinationWorkbook.Save("FormattedRowsOnly.xlsx");
    }
}