// Title: Unprotect a worksheet, add whole‑number data validation, and re‑protect it while preserving validation with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# unprotect worksheet, add whole number validation and protect again | How to keep data validation after protecting a sheet using Aspose.Cells for .NET | C# code to add validation to range A1:A10 after unprotecting a worksheet with Aspose.Cells | Preserve validation rules when re‑protecting an Excel worksheet with password in Aspose.Cells | Example of using Worksheet.Unprotect and Worksheet.Protect with validation in Aspose.Cells C# | Add whole number between 1 and 100 validation to Excel sheet using Aspose.Cells after unprotecting
// Tags: Worksheet.Unprotect and Protect with password Aspose.Cells | add whole number validation Aspose.Cells C# | preserve data validation after sheet protection Aspose.Cells | CellArea range validation Aspose.Cells | validation persistence Excel Aspose.Cells .NET | protect worksheet all options Aspose.Cells

using Aspose.Cells;
using System;

// This C# example demonstrates how to protect a worksheet with a password, unprotect it, apply a whole‑number data validation rule to cells A1:A10, re‑protect the sheet using the same password, verify that the validation remains, and save the workbook as UnprotectAddValidateProtect.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and rename it
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "DataSheet";

            // Protect the worksheet with a password (oldPassword is empty because the sheet is not yet protected)
            string password = "myPassword";
            worksheet.Protect(ProtectionType.All, password, string.Empty);

            // Unprotect the worksheet to allow modifications
            worksheet.Unprotect(password);

            // Define the range A1:A10 for whole‑number validation
            int firstRow = 0;      // Row 1 (zero‑based)
            int firstColumn = 0;   // Column A (zero‑based)
            int lastRow = 9;       // Row 10
            int lastColumn = 0;    // Column A

            // Create a CellArea for the target range
            CellArea area = new CellArea
            {
                StartRow = firstRow,
                StartColumn = firstColumn,
                EndRow = lastRow,
                EndColumn = lastColumn
            };

            // Add a data validation rule
            ValidationCollection validations = worksheet.Validations;
            int validationIndex = validations.Add(area);
            Validation validation = validations[validationIndex];
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "100";
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "Please enter a whole number between 1 and 100.";

            // Re‑protect the worksheet with the same password
            worksheet.Protect(ProtectionType.All, password, string.Empty);

            // Verify that the validation rule still exists after protection
            bool validationExists = worksheet.Validations.Count > 0;
            Console.WriteLine("Validation exists after protection: " + validationExists);

            // Save the workbook
            workbook.Save("UnprotectAddValidateProtect.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
