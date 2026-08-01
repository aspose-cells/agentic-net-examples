// Title: Check that a custom data‑validation formula on B1 points to A1 using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a custom validation to cell B1, sets its formula to "=A1" with A1 notation, retrieves the formula via GetFormula1, compares it case‑insensitively to "=A1", outputs the result, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | data validation | custom validation | set formula | GetFormula1 | A1 notation | workbook | validation reference | unit test
// Common Searches: Aspose.Cells set custom validation formula C# | GetFormula1 A1 notation example | verify validation references specific cell Aspose.Cells | how to check data validation formula in .NET
// Developer Intent: Confirm that the validation applied to B1 correctly references cell A1 on the first worksheet.
// Use Cases: Automated test to ensure a newly added validation rule points to the intended source cell. | Quality‑control script that validates formula integrity after workbook transformations. | Dynamic generation of validation rules with runtime verification of cell references.
// AI Prompts: Write C# code with Aspose.Cells that adds a custom validation to B1, sets the formula to "=A1", and asserts GetFormula1 returns "=A1". | Create a unit test in .NET using Aspose.Cells that verifies a validation's formula matches an expected A1 reference. | Provide a reusable method that applies a custom validation to any cell and confirms the formula points to a specified source cell.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationCheck
{
    // Creates a workbook, adds a custom validation to cell B1, sets its formula to "=A1" with A1 notation, retrieves the formula via GetFormula1, compares it case‑insensitively to "=A1", outputs the result, and saves the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define a validation that references cell A1
            // Apply the validation to cell B1 (row 0, column 1)
            CellArea area = new CellArea { StartRow = 0, StartColumn = 1, EndRow = 0, EndColumn = 1 };
            int validationIndex = worksheet.Validations.Add(area);
            Validation validation = worksheet.Validations[validationIndex];
            validation.Type = ValidationType.Custom;
            validation.Operator = OperatorType.Equal;
            // Set the formula to reference A1 using A1 notation
            validation.SetFormula1("=A1", false, false);

            // Retrieve the formula in A1 notation to verify the reference
            string formulaA1 = validation.GetFormula1(false, false);

            // Verify that the formula references the first worksheet cell A1
            bool referencesA1 = formulaA1.Equals("=A1", StringComparison.OrdinalIgnoreCase);
            Console.WriteLine($"Retrieved Formula: {formulaA1}");
            Console.WriteLine($"Does the formula reference A1? {referencesA1}");

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("ValidationReferenceCheck.xlsx");
        }
    }
}
