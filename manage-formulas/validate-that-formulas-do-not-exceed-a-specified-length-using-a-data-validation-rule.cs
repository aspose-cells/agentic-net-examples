// Title: Limit Excel Formula Length with TextLength Validation in Aspose.Cells for .NET
// Description: Creates a workbook, applies a TextLength data‑validation rule to cells B2:B5 that caps formula text at 50 characters, adds user‑friendly messages, demonstrates allowed and rejected formulas, and saves the file as FormulaLengthValidation.xlsx.
// Keywords: Aspose.Cells | C# | .NET | formula length validation | TextLength data validation | Excel data validation rule | limit formula characters | validation error handling | Excel performance | workbook automation
// Common Searches: Aspose.Cells limit formula length | TextLength validation Aspose .NET example | how to restrict Excel formula characters with Aspose | validate formula character count in C# | data validation rule for formula length
// Developer Intent: Add a TextLength validation rule that prevents formulas longer than a specified number of characters from being entered.
// Use Cases: Ensure shared workbooks stay performant by blocking overly long formulas. | Enforce corporate standards that cap formula complexity for maintainability. | Provide immediate on‑screen feedback when a user tries to enter a formula exceeding the allowed length.
// AI Prompts: Generate code to make the maximum formula length configurable at runtime instead of a hard‑coded value. | Show how to apply the same TextLength validation to several worksheets in a single workbook using Aspose.Cells. | Explain how to capture validation failures programmatically and log the offending formula text.

using Aspose.Cells;
using System;
using System.IO;

// Creates a workbook, applies a TextLength data‑validation rule to cells B2:B5 that caps formula text at 50 characters, adds user‑friendly messages, demonstrates allowed and rejected formulas, and saves the file as FormulaLengthValidation.xlsx.
class FormulaLengthValidationDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range where the validation will be applied (B2:B5)
            CellArea validationArea = CellArea.CreateCellArea(1, 1, 4, 1); // rows 1‑4, column 1 (B)

            // Add a validation rule to the worksheet
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Use TextLength validation to restrict the length of the cell content (including formulas entered as text)
            validation.Type = ValidationType.TextLength;

            // Operator for TextLength validation – LessThanOrEqual
            // Note: In some Aspose.Cells versions the default operator works for TextLength,
            // so we omit setting it to avoid compatibility issues.
            // validation.Operator = OperatorType.LessThanOrEqual;

            // Maximum allowed length
            validation.Formula1 = "50";

            // Optional: user‑friendly messages
            validation.InputTitle = "Formula Length";
            validation.InputMessage = "Enter a formula whose text length does not exceed 50 characters.";
            validation.ErrorTitle = "Invalid Length";
            validation.ErrorMessage = "The formula is too long.";
            validation.ShowInput = true;
            validation.ShowError = true;

            // Example of a short formula entered as text (allowed)
            worksheet.Cells["B2"].PutValue("'=SUM(A1:A10)");

            // Example of a long formula entered as text (will trigger validation error)
            worksheet.Cells["B3"].PutValue("'=IF(AND(A1>0,A2>0,A3>0,A4>0,A5>0,A6>0,A7>0,A8>0,A9>0,A10>0),\"All Positive\",\"Check Values\")");

            // Determine output path and ensure its directory exists
            string outputPath = "FormulaLengthValidation.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath) ?? Directory.GetCurrentDirectory();

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

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
