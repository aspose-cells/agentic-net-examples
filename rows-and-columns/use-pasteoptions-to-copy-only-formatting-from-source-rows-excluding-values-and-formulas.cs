using Aspose.Cells;

// Author: Aspose.Cells .NET example – copy only formatting from source rows

class Program
{
    static void Main()
    {
        // Load the source workbook containing the rows to copy
        Workbook sourceWorkbook = new Workbook("source.xlsx");
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Create (or load) the destination workbook where rows will be pasted
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Default copy options – no special behavior required
        CopyOptions copyOptions = new CopyOptions();

        // Paste options configured to copy only formats (no values or formulas)
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Formats   // Formats only
        };

        // Example: copy 5 rows starting from row 0 of the source sheet
        // and paste them beginning at row 10 of the destination sheet
        sourceSheet.Cells.CopyRows(
            sourceSheet.Cells,   // source cells
            0,                   // source row index (zero‑based)
            10,                  // destination row index (zero‑based)
            5,                   // number of rows to copy
            copyOptions,
            pasteOptions);

        // Save the resulting workbook
        destinationWorkbook.Save("result.xlsx");
    }
}