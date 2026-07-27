using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – copy rows with absolute references unchanged
class Program
{
    static void Main()
    {
        // Create a new workbook (replace with the provided create rule if available)
        Workbook workbook = new Workbook();

        // Source worksheet with formulas that use absolute references
        Worksheet srcSheet = workbook.Worksheets[0];
        srcSheet.Cells["A1"].Formula = "$B$1+10";
        srcSheet.Cells["A2"].Formula = "$B$2*2";

        // Destination worksheet
        Worksheet dstSheet = workbook.Worksheets.Add("Destination");

        // Configure copy and paste options
        CopyOptions copyOptions = new CopyOptions();

        PasteOptions pasteOptions = new PasteOptions
        {
            // Disable formula adjustment (relative reference shifting) by using No operation type
            OperationType = PasteOperationType.None,
            // Ensure formulas are copied
            PasteType = PasteType.Formulas
        };

        // Copy the first two rows from source to destination starting at row index 5 (row 6 in Excel)
        dstSheet.Cells.CopyRows(srcSheet.Cells, 0, 5, 2, copyOptions, pasteOptions);

        // Save the workbook (replace with the provided save rule if available)
        workbook.Save("RowsCopy_NoFormulaAdjustment.xlsx");
    }
}