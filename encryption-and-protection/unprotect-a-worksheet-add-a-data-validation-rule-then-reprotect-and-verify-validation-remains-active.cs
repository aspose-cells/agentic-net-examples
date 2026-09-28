// Title: How to unprotect an Excel worksheet, add whole‑number data validation to a cell, and re‑protect while preserving the validation using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that removes protection from a worksheet, inserts a whole‑number validation (1‑10) on cell A1, and then reapplies the same password protection. | Generate a snippet that checks the validation count after re‑protecting a sheet to confirm the rule remains active in Aspose.Cells. | Create a complete example that saves an .xlsx file after adding and preserving a data validation rule on a protected worksheet using Aspose.Cells for .NET.
// Common Searches: aspnet add integer range validation to a protected Excel sheet using Aspose.Cells | unprotect worksheet, apply data validation, protect again Aspose.Cells C# example | verify data validation persists after sheet protection Aspose.Cells .NET | how to check if validation exists after protecting an Excel worksheet with Aspose.Cells
// Tags: Aspose.Cells worksheet unprotect protect password | Aspose.Cells add whole number validation | Aspose.Cells preserve data validation after protection | Aspose.Cells save workbook as xlsx with validation | Aspose.Cells validation count verification

using System;
using Aspose.Cells;

namespace AsposeCellsDataValidationDemo
{
    // The example creates a new workbook, protects the first worksheet with a password, unprotects it, adds a whole‑number validation (values 1 to 10) to cell A1, re‑protects the sheet, confirms the validation rule still exists, and saves the file as DataValidationDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook with a default worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DemoSheet";

                // Define a password for protection
                string password = "SecurePwd123";

                // Protect the worksheet initially (oldPassword is empty for a new sheet)
                sheet.Protect(ProtectionType.All, password, string.Empty);

                // Unprotect the worksheet to allow modifications
                sheet.Unprotect(password);

                // Define the cell area for A1
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 0,
                    EndColumn = 0
                };

                // Add a data validation rule to cell A1 (integer between 1 and 10)
                int validationIndex = sheet.Validations.Add(area);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.WholeNumber;          // Whole number (integer) validation
                validation.Operator = OperatorType.Between;           // Between operator
                validation.Formula1 = "1";                            // Lower bound
                validation.Formula2 = "10";                           // Upper bound
                validation.InputMessage = "Please enter a whole number between 1 and 10.";
                validation.ErrorMessage = "The value entered is not within the allowed range.";

                // Re‑protect the worksheet with the same password
                sheet.Protect(ProtectionType.All, password, string.Empty);

                // Verify that the validation rule still exists after protection
                bool validationExists = sheet.Validations.Count > 0;
                Console.WriteLine($"Validation rule present after protection: {validationExists}");

                // Save the workbook to a file
                workbook.Save("DataValidationDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
