// Title: Copy Rows with List Data Validation Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add a list‑type data validation (Apple, Banana, Cherry) to a cell, copy the entire row with PasteOptions set to PasteType.Validation, confirm the validation appears in the destination cell, and save the workbook.
// Keywords: Aspose.Cells | C# | .NET | CopyRows | PasteOptions | PasteType.Validation | data validation | list validation | drop‑down list | Excel row copy | preserve validation | transfer rows | Excel automation
// Common Searches: Aspose.Cells copy row with data validation C# | how to preserve drop‑down list when copying rows in Excel using Aspose | PasteOptions Validation example Aspose.Cells | copy rows and keep list validation .NET | verify copied validation after row transfer Aspose
// Developer Intent: Copy a row that contains a list‑type data validation and ensure the validation rule is retained in the target row.
// Use Cases: Duplicate a template row that includes predefined drop‑down lists for new entries. | Programmatically add rows that share the same validation constraints across a worksheet. | Rearrange or insert rows while keeping all list validation rules intact.
// AI Prompts: Generate C# code with Aspose.Cells that copies multiple rows and preserves their data validation, formulas, and formatting. | Explain the role of PasteOptions.PasteType = PasteType.Validation and show how to programmatically confirm that validations were copied correctly. | Create a method that copies a range of rows, returns any validation mismatches, and logs the results.

using System;
using Aspose.Cells;

// Demonstrates how to add a list‑type data validation (Apple, Banana, Cherry) to a cell, copy the entire row with PasteOptions set to PasteType.Validation, confirm the validation appears in the destination cell, and save the workbook.
class TransferRowsWithValidation
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create source workbook and add a list validation
        // -------------------------------------------------
        Workbook srcWorkbook = new Workbook();
        Worksheet srcSheet = srcWorkbook.Worksheets[0];

        // Create a validation that provides a drop‑down list
        Validation srcValidation = srcSheet.Validations[srcSheet.Validations.Add()];
        srcValidation.Type = ValidationType.List;
        srcValidation.Formula1 = "Apple,Banana,Cherry";

        // Apply the validation to cell B1 (row 0, column 1)
        CellArea validationArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 1,
            EndRow = 0,
            EndColumn = 1
        };
        srcValidation.AddArea(validationArea);

        // -------------------------------------------------
        // 2. Copy the row that contains the validation
        // -------------------------------------------------
        // Prepare paste options to copy only validations
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Validation
        };

        // Copy row 0 (sourceRowIndex) to row 5 (destinationRowIndex)
        // rowNumber = 1 because we copy a single row
        srcSheet.Cells.CopyRows(
            srcSheet.Cells,   // source cells
            0,                // source row index
            5,                // destination row index
            1,                // number of rows to copy
            new CopyOptions(), // default copy options
            pasteOptions);     // paste options (validation only)

        // -------------------------------------------------
        // 3. Verify that the validation was copied correctly
        // -------------------------------------------------
        // The destination cell is B6 (row 5, column 1)
        Validation destValidation = srcSheet.Validations.GetValidationInCell(5, 1);

        Console.WriteLine("Destination validation present: " + (destValidation != null));
        if (destValidation != null)
        {
            Console.WriteLine("Validation type: " + destValidation.Type);
            Console.WriteLine("Validation list (Formula1): " + destValidation.Formula1);
        }

        // -------------------------------------------------
        // 4. Save the workbook to verify the result manually
        // -------------------------------------------------
        srcWorkbook.Save("TransferRowsWithValidation.xlsx");
    }
}
