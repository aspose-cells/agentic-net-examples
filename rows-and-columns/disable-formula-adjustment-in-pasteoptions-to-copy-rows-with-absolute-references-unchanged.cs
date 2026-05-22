using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data that will be referenced by absolute formulas
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(20);

        // Add a formula with absolute references in row 2 (zero‑based index 1)
        // The formula refers to $A$1 and $B$1, which should stay unchanged after copying
        sheet.Cells["C2"].Formula = "=$A$1+$B$1";

        // Prepare copy and paste options
        // CopyOptions – default behavior (no special handling required)
        CopyOptions copyOptions = new CopyOptions();

        // PasteOptions – copy formulas only and do not apply any operation type
        // This prevents Excel‑like relative adjustment of formulas during the copy
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Formulas,          // copy formulas as they are
            OperationType = PasteOperationType.None  // no arithmetic operation on the copied data
        };

        // Copy the entire row 2 (index 1) to row 4 (index 3)
        // Parameters: source cells, source row index, destination row index, number of rows, copy options, paste options
        sheet.Cells.CopyRows(sheet.Cells, 1, 3, 1, copyOptions, pasteOptions);

        // Save the workbook to verify that the formula in C4 remains "=$A$1+$B$1"
        workbook.Save("AbsoluteReferenceCopy_Output.xlsx");
    }
}