// Title: Create a custom data validation rule in Aspose.Cells for .NET that limits formula length to 50 characters in cells A1:A100
// AI Prompts: Generate C# code using Aspose.Cells to add a custom validation that restricts any formula entered in cells A1 through A100 to a maximum of 50 characters, displaying an error dialog when exceeded. | Write an Aspose.Cells example that builds a validation formula with LEN and FORMULATEXT to enforce a formula‑length constraint and saves the workbook as an .xlsx file. | Provide a step‑by‑step C# snippet that creates a workbook, defines a CellArea for A1:A100, sets Validation.Type to Custom, and uses validation.Formula1 = "LEN(FORMULATEXT(A1))<=50".
// Common Searches: aspnet aspose.cells custom validation to limit formula characters | how to enforce maximum formula length in Excel using Aspose.Cells C# | C# Aspose.Cells LEN FORMULATEXT validation example | restrict formula size for a range with Aspose.Cells data validation | error message for formula length violation Aspose.Cells
// Tags: custom validation formula length Aspose.Cells | LEN FORMULATEXT constraint C# | apply data validation to range A1:A100 Aspose | formula character limit workbook Aspose.Cells | error alert for formula length Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, defines a custom validation on cells A1:A100 that uses LEN(FORMULATEXT(...)) to ensure formulas do not exceed 50 characters, shows an error message when the limit is breached, and saves the file as FormulaLengthValidation.xlsx.
class FormulaLengthValidation
{
    static void Main()
    {
        try
        {
            // Define the maximum allowed length for formulas
            int maxFormulaLength = 50;

            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range where the validation will be applied (e.g., A1:A100)
            CellArea validationRange = new CellArea
            {
                StartRow = 0,    // Row 1 (zero‑based index)
                EndRow = 99,     // Row 100
                StartColumn = 0, // Column A
                EndColumn = 0    // Column A
            };

            // Add a custom data validation rule to the specified range (lifecycle rule: create)
            int validationIndex = sheet.Validations.Add(validationRange);
            Validation validation = sheet.Validations[validationIndex];

            // Set the validation type to Custom
            validation.Type = ValidationType.Custom;

            // Get the address of the top‑left cell in the range (e.g., "A1")
            string topLeftCellAddress = sheet.Cells[validationRange.StartRow, validationRange.StartColumn].Name;

            // Build the custom formula using FORMULATEXT and LEN
            validation.Formula1 = $"LEN(FORMULATEXT({topLeftCellAddress}))<={maxFormulaLength}";

            // Optional: display an error message when the rule is violated
            validation.ShowError = true;
            validation.ErrorTitle = "Formula Too Long";
            validation.ErrorMessage = $"Formulas must not exceed {maxFormulaLength} characters.";

            // Save the workbook (lifecycle rule: save)
            workbook.Save("FormulaLengthValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
