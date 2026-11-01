// Title: Add a custom data validation rule in Aspose.Cells for .NET to limit Excel formula length
// Common Searches: how to enforce maximum formula length using Aspose.Cells C# | Aspose.Cells custom validation rule to restrict formula characters | C# code to validate Excel formula length with LEN function in Aspose.Cells | set error message for formula length violation in Aspose.Cells workbook | apply data validation to range B2:B10 for formula length in .NET
// Tags: Aspose.Cells validation rule API | restrict Excel formula characters programmatically | LEN function based data validation Aspose | set formula length limit in .NET workbook | configure validation error message Aspose.Cells | target cell area B2:B10 for validation

using Aspose.Cells;
using System;

// The example creates a workbook, defines the range B2:B10, and adds a custom data validation rule that uses the LEN function to ensure any entered formula does not exceed a specified character count (e.g., 50). It also configures an error title and message and saves the file as FormulaLengthValidation.xlsx.
class Program
{
    static void Main()
    {
        // Maximum allowed length for a formula
        int maxLength = 50;

        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range where the validation will be applied (e.g., B2:B10)
        CellArea validationArea = new CellArea
        {
            StartRow = 1,      // Row index is zero‑based (B2)
            StartColumn = 1,   // Column index is zero‑based (B)
            EndRow = 9,        // B10
            EndColumn = 1
        };

        // Add a data validation rule to the specified range
        Validation validation = sheet.Validations[sheet.Validations.Add(validationArea)];
        validation.Type = ValidationType.Custom;

        // Use the LEN function to ensure the formula length does not exceed maxLength.
        // The formula refers to the cell being validated using a relative reference.
        validation.Formula1 = $"LEN({sheet.Cells[validationArea.StartRow, validationArea.StartColumn].Name})<={maxLength}";

        // Configure the error message shown when the rule is violated
        validation.ShowError = true;
        validation.ErrorTitle = "Formula Too Long";
        validation.ErrorMessage = $"The formula must not exceed {maxLength} characters.";

        // Save the workbook
        workbook.Save("FormulaLengthValidation.xlsx");
    }
}
