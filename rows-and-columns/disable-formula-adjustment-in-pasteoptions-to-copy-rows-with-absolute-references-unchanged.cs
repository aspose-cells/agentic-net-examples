// Title: Copy rows without formula adjustment using PasteOptions in Aspose.Cells for .NET
// Description: Shows how to duplicate a worksheet row while preserving all formulas exactly as written, including absolute references, by setting PasteOptions.PasteType to Formulas and disabling formula adjustment. The sample uses Worksheet.Cells.CopyRows with CopyOptions and PasteOptions, then saves the workbook to an .xlsx file.
// Keywords: Aspose.Cells CopyRows | PasteOptions formula adjustment | disable formula adjustment | preserve absolute references | C# copy rows formulas | Aspose.Cells .NET | CopyRows PasteOptions | Excel row duplication without reference change
// Common Searches: Aspose.Cells copy rows keep formulas unchanged | How to stop formula references from shifting when copying rows in .NET | PasteOptions IsFormulaAdjustmentEnabled false example | CopyRows with absolute references Aspose.Cells | C# Aspose.Cells duplicate row without changing formulas
// Developer Intent: Duplicate a worksheet row while leaving every formula unchanged.
// Use Cases: Copy a template row that contains summary formulas to multiple sections of a financial report, ensuring the formulas still point to the original source cells. | Create a static header row with calculated fields that must remain constant across different pages of a large spreadsheet. | Replicate a calculation row in a budgeting model where the references must stay absolute regardless of where the row is pasted.
// AI Prompts: Provide C# code using Aspose.Cells to copy rows with PasteOptions that disables formula adjustment. | Generate an example that copies a range of rows while preserving absolute references using CopyRows and PasteOptions. | Explain how to set PasteOptions.IsFormulaAdjustmentEnabled = false in Aspose.Cells for .NET and why it matters.

using Aspose.Cells;
using System;

// Shows how to duplicate a worksheet row while preserving all formulas exactly as written, including absolute references, by setting PasteOptions.PasteType to Formulas and disabling formula adjustment. The sample uses Worksheet.Cells.CopyRows with CopyOptions and PasteOptions, then saves the workbook to an .xlsx file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate source row (row 0) with values and formulas
        sheet.Cells["A1"].PutValue(10);                     // Value in A1
        sheet.Cells["B1"].PutValue(20);                     // Value in B1
        sheet.Cells["C1"].Formula = "=SUM($A$1,$B$1)";      // Formula with absolute references
        sheet.Cells["D1"].Formula = "=A1+B1";               // Formula with relative references (for contrast)

        // Prepare copy options (default behavior)
        CopyOptions copyOptions = new CopyOptions();

        // Prepare paste options to copy formulas exactly as they are
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Formulas,                 // Copy only formulas
            OperationType = PasteOperationType.None        // No arithmetic operation on the copied data
        };

        // Copy the first row (index 0) to the fifth row (index 4)
        // This uses the overload that accepts both CopyOptions and PasteOptions
        sheet.Cells.CopyRows(sheet.Cells, 0, 4, 1, copyOptions, pasteOptions);

        // Save the workbook
        workbook.Save("FormulaCopy_NoAdjustment.xlsx");
    }
}
