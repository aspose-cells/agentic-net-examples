// Title: Copy Rows with List Data Validation and Verify Rules Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a list‑type data validation, copy rows with Aspose.Cells' CopyRows method, preserve the validation via PasteOptions.Validation, and confirm the copied rule by retrieving the validation object and its first list item. The workbook is saved as ValidationCopyResult.xlsx.
// Keywords: Aspose.Cells | CopyRows | PasteOptions.Validation | C# | .NET | Excel data validation | list validation copy | preserve validation | duplicate rows | Excel automation
// Common Searches: Aspose.Cells copy rows keep validation | Copy rows with dropdown list using C# | PasteOptions.Validation example | Verify data validation after copying rows | How to duplicate rows with list validation in Aspose.Cells
// Developer Intent: Copy rows that contain list‑type data validation and ensure the validation rules stay functional in the destination range.
// Use Cases: Replicate a template section that includes dropdown lists to a new area for additional entries. | Move a block of rows with list validations to a summary section while keeping the original source range reference. | Programmatically confirm that copied cells retain the List validation type and return the expected list values.
// AI Prompts: Generate C# code with Aspose.Cells to copy rows 0‑2 to row 10 and preserve all data validation rules. | Show how to compare list validation values before and after copying rows using Aspose.Cells. | Explain the impact of setting PasteOptions.PasteType = PasteType.Validation when using the CopyRows method.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationCopyDemo
{
    // Demonstrates how to create a list‑type data validation, copy rows with Aspose.Cells' CopyRows method, preserve the validation via PasteOptions.Validation, and confirm the copied rule by retrieving the validation object and its first list item. The workbook is saved as ValidationCopyResult.xlsx.
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook ----------
            Workbook srcWb = new Workbook();
            Worksheet srcSheet = srcWb.Worksheets[0];

            // Populate list source range A1:A4
            srcSheet.Cells["A1"].PutValue("Apple");
            srcSheet.Cells["A2"].PutValue("Banana");
            srcSheet.Cells["A3"].PutValue("Cherry");
            srcSheet.Cells["A4"].PutValue("Date");

            // Add validation of type List to cells B1:B3 (rows 0‑2, column 1)
            ValidationCollection validations = srcSheet.Validations;
            // Create a validation and apply to the first cell (B1)
            Validation val = validations[validations.Add()];
            val.Type = ValidationType.List;
            val.Formula1 = "A1:A4"; // reference the list range
            // Apply the same validation to B2 and B3
            val.AddArea(CellArea.CreateCellArea(1, 1, 1, 1));
            val.AddArea(CellArea.CreateCellArea(2, 1, 2, 1));

            // Put some sample data in the rows to be copied
            srcSheet.Cells["B1"].PutValue("Apple");
            srcSheet.Cells["B2"].PutValue("Banana");
            srcSheet.Cells["B3"].PutValue("Cherry");

            // ---------- Copy rows (including validations) ----------
            // Prepare copy options (default)
            CopyOptions copyOptions = new CopyOptions();

            // Prepare paste options to include validations
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.Validation   // ensure validations are copied
            };

            // Copy rows 0‑2 (sourceRowIndex = 0, rowNumber = 3) to destination starting at row 5
            srcSheet.Cells.CopyRows(srcSheet.Cells, 0, 5, 3, copyOptions, pasteOptions);

            // ---------- Verify copied validation ----------
            // Get validation applied to destination cell B6 (row index 5, column 1)
            Validation copiedVal = srcSheet.Validations.GetValidationInCell(5, 1);
            if (copiedVal != null && copiedVal.Type == ValidationType.List)
            {
                Console.WriteLine("Validation copied successfully. Type: " + copiedVal.Type);
                // Retrieve list values for the first item in the list
                object listValue = copiedVal.GetListValue(0, 0);
                Console.WriteLine("First list value after copy: " + listValue);
            }
            else
            {
                Console.WriteLine("Validation was not copied correctly.");
            }

            // ---------- Save workbook (optional) ----------
            srcWb.Save("ValidationCopyResult.xlsx");
        }
    }
}
