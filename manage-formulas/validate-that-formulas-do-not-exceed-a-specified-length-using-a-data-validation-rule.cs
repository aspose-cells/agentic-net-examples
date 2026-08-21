// Title: Limit Excel Formula Length with Custom Data Validation in Aspose.Cells for .NET
// Description: Shows how to apply a custom validation rule (LEN(FORMULATEXT(A1))<=50) to cells A1:A10 using Aspose.Cells for C#. Includes user prompts, error alerts, and workbook saving.
// Keywords: Aspose.Cells | C# data validation | custom validation rule | formula length limit | LEN FORMULATEXT | Excel formula character count | restrict formula size | validation rule .NET
// Common Searches: Aspose.Cells limit formula characters | C# data validation formula length | How to use FORMULATEXT in Aspose.Cells | Set custom validation for formula size in Excel via code | Validate Excel formula length programmatically
// Developer Intent: Create a data‑validation rule that blocks formulas exceeding a defined character count.
// Use Cases: Enforce a maximum formula length in a column where users enter calculations, keeping formulas under 50 characters. | Prevent overly complex formulas in shared workbooks to maintain performance and readability. | Apply corporate standards that cap formula size before the workbook is saved or distributed.
// AI Prompts: Generate C# code with Aspose.Cells that adds a custom validation limiting formulas to 100 characters for range B2:B20. | Explain how FORMULATEXT and LEN can be combined in an Aspose.Cells Validation to enforce a maximum formula size. | Provide an example of handling validation errors when a formula exceeds the allowed length in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationExample
{
    // Shows how to apply a custom validation rule (LEN(FORMULATEXT(A1))<=50) to cells A1:A10 using Aspose.Cells for C#. Includes user prompts, error alerts, and workbook saving.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range where the validation will be applied (A1:A10)
            CellArea validationArea = CellArea.CreateCellArea(0, 0, 9, 0); // rows 0‑9, column 0 (A)

            // Add a new validation to the worksheet
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Set validation type to Custom and no operator (the formula itself determines validity)
            validation.Type = ValidationType.Custom;
            validation.Operator = OperatorType.None;

            // Custom formula: ensure the length of the cell's formula does not exceed 50 characters.
            // FORMULATEXT returns the formula as a string; LEN measures its length.
            // The relative reference A1 will adjust for each cell in the area.
            validation.Formula1 = "LEN(FORMULATEXT(A1))<=50";

            // Optional: provide user-friendly messages
            validation.InputTitle = "Formula Length Check";
            validation.InputMessage = "Enter a formula no longer than 50 characters.";
            validation.ErrorTitle = "Invalid Formula Length";
            validation.ErrorMessage = "The formula exceeds the allowed length of 50 characters.";
            validation.ShowInput = true;
            validation.ShowError = true;

            // Save the workbook
            workbook.Save("FormulaLengthValidation.xlsx");
        }
    }
}
