// Title: C# – Unprotect Worksheet, Add List Validation, Re‑protect and Verify with Aspose.Cells
// Description: Demonstrates how to temporarily unprotect a worksheet, insert a list‑type data validation (e.g., Option1, Option2, Option3) into cell A1, re‑apply the same password protection, and confirm that the validation remains active, using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# worksheet protect | unprotect worksheet Aspose.Cells | add data validation list Aspose.Cells | re‑protect worksheet validation persists | .NET Excel protection example | Aspose.Cells validation after protect | C# Excel dropdown list protected sheet
// Common Searches: Aspose.Cells unprotect worksheet then add validation | C# add dropdown list to protected Excel sheet | verify data validation after protecting worksheet Aspose.Cells | how to protect worksheet after adding validation in .NET | Aspose.Cells example for worksheet protection and validation
// Developer Intent: Temporarily lift worksheet protection, add a list‑type data validation to a cell, re‑apply protection, and ensure the validation is still functional.
// Use Cases: Updating a locked template by inserting new dropdown lists before distribution. | Programmatically enforcing data entry rules in a workbook that must stay protected for end users. | Generating reports where validation settings must survive worksheet protection to maintain data integrity.
// AI Prompts: Write C# code with Aspose.Cells that unprotects a sheet, adds a list validation to cell B2, re‑protects with a password, and checks the validation persists. | Explain how Aspose.Cells handles data validation objects when a worksheet is protected and how to retrieve them after re‑protection. | Provide troubleshooting steps when a data validation disappears after re‑protecting a worksheet using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationDemo
{
    // Demonstrates how to temporarily unprotect a worksheet, insert a list‑type data validation (e.g., Option1, Option2, Option3) into cell A1, re‑apply the same password protection, and confirm that the validation remains active, using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // 1. Protect the worksheet with a password (initial protection)
            // -----------------------------------------------------------------
            string password = "pwd123";
            worksheet.Protect(ProtectionType.All, password, null);
            Console.WriteLine("Worksheet initially protected: " + worksheet.IsProtected);

            // -----------------------------------------------------------------
            // 2. Unprotect the worksheet using the password
            // -----------------------------------------------------------------
            worksheet.Unprotect(password);
            Console.WriteLine("Worksheet after unprotect: " + !worksheet.IsProtected);

            // -----------------------------------------------------------------
            // 3. Add a data validation rule to cell A1 (row 0, column 0)
            // -----------------------------------------------------------------
            // Define the cell area for the validation (single cell A1)
            CellArea validationArea = CellArea.CreateCellArea(0, 0, 0, 0);

            // Add the validation to the worksheet's validation collection
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation as a list with three options
            validation.Type = ValidationType.List;
            validation.Formula1 = "Option1,Option2,Option3";
            validation.InCellDropDown = true;
            validation.ShowInput = true;
            validation.InputMessage = "Select an option from the list.";
            validation.ErrorMessage = "Invalid selection.";
            validation.ShowError = true;

            Console.WriteLine("Validation added to cell A1.");

            // -----------------------------------------------------------------
            // 4. Re‑protect the worksheet with the same password
            // -----------------------------------------------------------------
            worksheet.Protect(ProtectionType.All, password, null);
            Console.WriteLine("Worksheet re‑protected: " + worksheet.IsProtected);

            // -----------------------------------------------------------------
            // 5. Verify that the validation still exists after protection
            // -----------------------------------------------------------------
            Validation retrievedValidation = worksheet.Validations.GetValidationInCell(0, 0);
            if (retrievedValidation != null)
            {
                Console.WriteLine("Validation type after re‑protect: " + retrievedValidation.Type);
            }
            else
            {
                Console.WriteLine("No validation found after re‑protect.");
            }

            // -----------------------------------------------------------------
            // 6. Save the workbook
            // -----------------------------------------------------------------
            workbook.Save("UnprotectAddValidateProtect.xlsx");
            Console.WriteLine("Workbook saved as UnprotectAddValidateProtect.xlsx");
        }
    }
}
