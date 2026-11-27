// Title: Add a custom data validation rule in Aspose.Cells for .NET to restrict Excel formula length
// Common Searches: Aspose.Cells C# how to limit formula character count with data validation | create custom validation in .NET workbook to enforce maximum formula length | using LEN function in Aspose.Cells validation to check formula size | prevent Excel formulas longer than 20 characters in Aspose.Cells | apply formula length restriction to a range with Aspose.Cells API
// Tags: custom data validation formula length Aspose.Cells | maximum formula characters C# Aspose.Cells | LEN function validation rule .NET | restrict Excel formula size Aspose.Cells | validation collection add custom rule C#

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, writes sample formulas, defines a maximum formula length (e.g., 20 characters), and adds a custom data validation rule to cells A1:A100 using the LEN function (LEN(A1)<=20). The rule shows an error message when a formula exceeds the limit, then saves the file as FormulaLengthValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example formulas placed in column A
            sheet.Cells["A1"].Formula = "SUM(B1:B10)";
            sheet.Cells["A2"].Formula = "IF(C1>0, D1, E1)";

            // Define the maximum allowed formula length
            int maxLength = 20;

            // Apply data validation to the range A1:A100
            int firstRow = 0;   // zero‑based index for row 1
            int lastRow = 99;   // row 100

            // Get the validation collection for the worksheet
            ValidationCollection validations = sheet.Validations;

            // Define the address of the range to validate (for reference only)
            string rangeAddress = $"A{firstRow + 1}:A{lastRow + 1}";

            // Convert the address to a CellArea (required by the Add method)
            // Column A corresponds to column index 0
            CellArea area = new CellArea
            {
                StartRow = firstRow,
                StartColumn = 0,
                EndRow = lastRow,
                EndColumn = 0
            };

            // Add a custom validation rule for the specified range
            int validationIndex = validations.Add(area);
            Validation validation = validations[validationIndex];
            validation.Type = ValidationType.Custom;

            // Custom formula: length of the cell's formula must be <= maxLength
            // Excel evaluates the formula relative to each cell in the range.
            validation.Formula1 = $"LEN(A1)<={maxLength}";

            // Configure error display
            validation.ShowError = true;
            validation.ErrorTitle = "Formula Too Long";
            validation.ErrorMessage = $"The formula length must not exceed {maxLength} characters.";

            // Define output file path
            string outputPath = "FormulaLengthValidation.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
