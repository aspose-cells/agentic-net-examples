// Title: Copy rows with list‑type data validation using Aspose.Cells for .NET
// Description: Demonstrates how to create a list validation (dropdown) that references cells A1:A4, apply it to B1‑B3, copy rows 0‑2 to rows 5‑7 with CopyRows, and verify that the validation is preserved in B6 using PasteOptions.PasteType = Validation and GetListValue.
// Keywords: Aspose.Cells copy rows validation | preserve dropdown list Aspose.Cells | CopyRows PasteOptions Validation .NET | verify data validation after copy | GetListValue Aspose.Cells example
// Common Searches: Aspose.Cells copy rows with validation | how to keep dropdown list when copying rows in .NET | validate copied data validation in Aspose.Cells | CopyRows preserving data validation
// Developer Intent: Programmatically duplicate rows that contain list‑type data validation and confirm that the validation rules remain active in the target rows.
// Use Cases: Clone a template section that includes dropdown lists to generate multiple entry blocks while retaining functional dropdowns. | Automate the replication of rows with validation to a new area of a worksheet and ensure the copied cells still reference the original list range. | Extract and display the items of a copied list validation to verify that the validation was transferred correctly.
// AI Prompts: Show me a C# example that copies rows with list validation using Aspose.Cells and checks the validation after the copy. | How can I update the validation formula when the source list range moves after copying rows in Aspose.Cells? | Explain how GetListValue works on a copied validation and how to enumerate its items.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationCopyDemo
{
    // Demonstrates how to create a list validation (dropdown) that references cells A1:A4, apply it to B1‑B3, copy rows 0‑2 to rows 5‑7 with CopyRows, and verify that the validation is preserved in B6 using PasteOptions.PasteType = Validation and GetListValue.
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook ----------
            Workbook srcWorkbook = new Workbook();
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Populate list source range A1:A4
            srcSheet.Cells["A1"].PutValue("Apple");
            srcSheet.Cells["A2"].PutValue("Banana");
            srcSheet.Cells["A3"].PutValue("Cherry");
            srcSheet.Cells["A4"].PutValue("Date");

            // Add a validation of type List to cell B1 (row 0, column 1)
            Validation validation = srcSheet.Validations[srcSheet.Validations.Add()];
            validation.Type = ValidationType.List;
            validation.Formula1 = "A1:A4";               // reference the list range
            validation.AddArea(CellArea.CreateCellArea(0, 1, 0, 1)); // apply to B1

            // Also apply the same validation to B2 and B3 (rows 1‑2)
            validation.AddArea(CellArea.CreateCellArea(1, 1, 2, 1));

            // Fill some data in the rows that will be copied
            srcSheet.Cells["A1"].PutValue("Row0");
            srcSheet.Cells["B1"].PutValue("Apple");
            srcSheet.Cells["A2"].PutValue("Row1");
            srcSheet.Cells["B2"].PutValue("Banana");
            srcSheet.Cells["A3"].PutValue("Row2");
            srcSheet.Cells["B3"].PutValue("Cherry");

            // ---------- Copy rows 0‑2 to rows 5‑7 ----------
            // Use CopyOptions (default) and PasteOptions with Validation paste type
            CopyOptions copyOptions = new CopyOptions(); // default options
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.Validation   // ensure validations are copied
            };

            // Copy three rows starting at source index 0 to destination index 5
            srcSheet.Cells.CopyRows(srcSheet.Cells, 0, 5, 3, copyOptions, pasteOptions);

            // ---------- Verify copied validation ----------
            // The validation should now exist in cell B6 (row 5, column 1)
            Validation copiedValidation = srcSheet.Validations.GetValidationInCell(5, 1);
            if (copiedValidation != null && copiedValidation.Type == ValidationType.List)
            {
                Console.WriteLine("Validation successfully copied to B6.");
                // Retrieve list values for the copied cell
                object listValue = copiedValidation.GetListValue(5, 1);
                // When the list is a range reference, GetListValue returns a ReferredArea.
                // For demonstration, we will read the actual cell values from the referenced range.
                if (listValue is ReferredArea area)
                {
                    Console.WriteLine("List values from referenced range:");
                    for (int r = area.StartRow; r <= area.EndRow; r++)
                    {
                        for (int c = area.StartColumn; c <= area.EndColumn; c++)
                        {
                            Console.WriteLine($" - {srcSheet.Cells[r, c].StringValue}");
                        }
                    }
                }
                else if (listValue is object[] arr)
                {
                    Console.WriteLine("List values (comma‑separated):");
                    foreach (var v in arr)
                        Console.WriteLine($" - {v}");
                }
                else
                {
                    Console.WriteLine($"List value: {listValue}");
                }
            }
            else
            {
                Console.WriteLine("Validation was not copied correctly.");
            }

            // ---------- Save the workbook (optional) ----------
            srcWorkbook.Save("ValidationCopyResult.xlsx");
        }
    }
}
