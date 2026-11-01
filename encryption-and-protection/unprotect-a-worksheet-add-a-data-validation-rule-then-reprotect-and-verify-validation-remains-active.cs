// Title: C# example: Unprotect a worksheet, add whole‑number data validation (1‑100) to A1:A10, then re‑protect while preserving the validation using Aspose.Cells
// Common Searches: C# Aspose.Cells how to unprotect a worksheet, add numeric validation and protect it again | preserve data validation after protecting an Excel sheet with Aspose.Cells for .NET | add whole number validation to a range in an Excel workbook using Aspose.Cells C# | verify that data validation rules remain after re‑protecting a worksheet in Aspose.Cells | Aspose.Cells protect worksheet with password then unprotect to modify validations
// Tags: worksheet protection password Aspose.Cells | data validation whole number range Aspose.Cells | retain validation after re‑protecting sheet | validation collection usage Aspose.Cells | save Excel file with validation .NET

using Aspose.Cells;
using System;
using System.IO;

// The code creates a new workbook, protects the first worksheet with a password, unprotects it, adds a whole‑number validation (1‑100) to cells A1:A10, re‑protects the sheet, confirms the validation still exists, and saves the file as DataValidationDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Protect the worksheet with a password (oldPassword not required for initial protect)
            sheet.Protect(ProtectionType.All, "myPassword", null);

            // Unprotect the worksheet to allow modifications
            sheet.Unprotect("myPassword");

            // Add a data validation rule: whole numbers between 1 and 100 in A1:A10
            ValidationCollection validations = sheet.Validations;
            // Create a CellArea for the range A1:A10
            CellArea range = CellArea.CreateCellArea("A1", "A10");
            int index = validations.Add(range); // add validation for the range
            Validation validation = validations[index];
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "100";
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid entry";
            validation.ErrorMessage = "Please enter a whole number between 1 and 100.";

            // Re‑protect the worksheet
            sheet.Protect(ProtectionType.All, "myPassword", null);

            // Verify that the validation rule is still present
            bool validationExists = sheet.Validations.Count > 0;
            Console.WriteLine("Validation rule present after re‑protect: " + validationExists);

            // Save the workbook
            string outputPath = "DataValidationDemo.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
