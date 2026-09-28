// Title: Copy rows that contain list‑type data validation and keep the validation functional using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells Cells.CopyRows with PasteOptions.PasteType.All to copy a range of rows from a source worksheet to a destination worksheet while preserving list‑type data validation. | After the copy, retrieve the validation object in the destination cell and confirm that its Type is List and its Formula1 still references the original range. | Save the destination workbook to an .xlsx file after the row transfer and validation verification.
// Common Searches: aspnet copy rows with dropdown validation using Aspose.Cells | preserve list validation when copying rows between worksheets in C# | how to verify data validation after copying rows with Aspose.Cells | Aspose.Cells Cells.CopyRows example with PasteOptions.All | transfer rows containing data validation to another workbook .NET
// Tags: copy rows with list validation Aspose.Cells | preserve data validation during row copy C# | Cells.CopyRows PasteOptions.All | validate transferred dropdown list Aspose.Cells | save workbook after row transfer Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a source workbook with a list validation in B1 referencing A1:A4, copies the first four rows to a new workbook using Cells.CopyRows with PasteOptions.PasteType.All, checks that the validation was transferred unchanged, and saves the result as RowCopyWithValidation.xlsx.
class TransferRowsWithValidation
{
    static void Main()
    {
        // ---------- Create source workbook with a validation list ----------
        Workbook srcWb = new Workbook();
        Worksheet srcSheet = srcWb.Worksheets[0];

        // Populate the list source values (A1:A4)
        srcSheet.Cells["A1"].PutValue("Apple");
        srcSheet.Cells["A2"].PutValue("Banana");
        srcSheet.Cells["A3"].PutValue("Cherry");
        srcSheet.Cells["A4"].PutValue("Date");

        // Add a List‑type validation to cell B1 that refers to A1:A4
        Validation srcValidation = srcSheet.Validations[srcSheet.Validations.Add()];
        srcValidation.Type = ValidationType.List;
        srcValidation.Formula1 = "A1:A4";
        srcValidation.AddArea(CellArea.CreateCellArea(0, 1, 0, 1)); // B1

        // ---------- Create destination workbook ----------
        Workbook destWb = new Workbook();
        Worksheet destSheet = destWb.Worksheets[0];

        // ---------- Copy rows (including the list source rows) ----------
        // Copy rows 0‑3 from source to destination (row 0 in destination)
        // Use CopyOptions (default) and PasteOptions with PasteType.All to copy
        // data, formats, and validations.
        CopyOptions copyOptions = new CopyOptions();
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.All   // ensures validations are copied
        };

        destSheet.Cells.CopyRows(
            srcSheet.Cells,   // source cells
            0,                // source start row
            0,                // destination start row
            4,                // number of rows to copy
            copyOptions,
            pasteOptions);

        // ---------- Verify that the validation was copied ----------
        Validation destValidation = destSheet.Validations.GetValidationInCell(0, 1); // B1 in destination
        if (destValidation != null &&
            destValidation.Type == ValidationType.List &&
            destValidation.Formula1 == "A1:A4")
        {
            Console.WriteLine("Validation copied successfully and remains functional.");
        }
        else
        {
            Console.WriteLine("Validation copy failed.");
        }

        // ---------- Save the destination workbook ----------
        destWb.Save("RowCopyWithValidation.xlsx");
    }
}
