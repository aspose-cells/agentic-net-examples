using System;
using Aspose.Cells;

namespace AsposeCellsColumnCopyDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook and populate a column ----------
            Workbook srcWorkbook = new Workbook();
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Set column width (in characters) for column A
            srcSheet.Cells.SetColumnWidth(0, 25); // 25 characters wide

            // Populate different data types in column A
            srcSheet.Cells["A1"].PutValue("Text");                     // string
            srcSheet.Cells["A2"].PutValue(12345);                     // integer
            srcSheet.Cells["A3"].PutValue(123.456);                   // double
            srcSheet.Cells["A4"].PutValue(DateTime.Now);              // DateTime
            srcSheet.Cells["A5"].PutValue(true);                      // boolean

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // ---------- Prepare paste options ----------
            PasteOptions pasteOptions = new PasteOptions
            {
                // Copy everything (values, formats, column widths, etc.)
                PasteType = PasteType.All
            };

            // ---------- Copy the column ----------
            // Parameters: sourceCells, sourceColumnIndex, destinationColumnIndex, columnNumber, pasteOptions
            destSheet.Cells.CopyColumns(
                srcSheet.Cells,   // source cells
                0,                // source column index (A)
                0,                // destination column index (A)
                1,                // number of columns to copy
                pasteOptions);   // paste options

            // ---------- Save the destination workbook ----------
            destWorkbook.Save("ColumnCopyResult.xlsx");
        }
    }
}