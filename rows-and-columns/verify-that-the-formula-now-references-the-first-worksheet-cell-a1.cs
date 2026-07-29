// Title: C# – Verify a data‑validation formula on B1 points to A1 using Aspose.Cells
// Description: Creates a workbook, adds a decimal validation to cell B1, sets Formula1 to "=A1" in A1 notation, reads it back with GetFormula1, checks case‑insensitively whether it equals "A1" or "=A1", prints the outcome and saves the file.
// Keywords: Aspose.Cells C# validation formula | SetFormula1 A1 notation | GetFormula1 verification | data validation cell reference | Aspose.Cells workbook example | C# Excel validation check | reference A1 in validation | Aspose.Cells API
// Common Searches: Aspose.Cells set validation formula to A1 | How to get validation formula in Aspose.Cells C# | Check if validation rule references a specific cell Aspose.Cells | C# verify data validation cell reference Aspose.Cells | Aspose.Cells GetFormula1 returns A1
// Developer Intent: Confirm that the validation rule applied to cell B1 correctly references cell A1.
// Use Cases: Ensure a numeric validation uses a threshold stored in A1 before applying it to other cells. | Audit existing workbooks to verify that all data‑validation formulas point to the intended source cell. | Generate a compliance report listing each validation rule and its referenced cell.
// AI Prompts: Write C# code with Aspose.Cells that sets a validation formula to reference cell A1 and then verifies the reference. | Explain how GetFormula1 returns the formula string and how to perform a case‑insensitive comparison to confirm the reference. | Show how to iterate through all validations in a worksheet and flag any that do not reference the expected cell.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationCheck
{
    // Creates a workbook, adds a decimal validation to cell B1, sets Formula1 to "=A1" in A1 notation, reads it back with GetFormula1, checks case‑insensitively whether it equals "A1" or "=A1", prints the outcome and saves the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the area where the validation will be applied (cell B1)
            CellArea validationArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 1,   // Column B
                EndRow = 0,
                EndColumn = 1
            };

            // Add a validation to the worksheet and obtain its index
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation type (e.g., Decimal) and operator
            validation.Type = ValidationType.Decimal;
            validation.Operator = OperatorType.GreaterOrEqual;

            // Set the validation formula to reference the first cell A1 using A1 notation
            // Parameters: formula string, isR1C1 = false (A1 style), isLocal = false
            validation.SetFormula1("=A1", false, false);

            // Retrieve the formula back in A1 notation to verify the reference
            // Parameters: isR1C1 = false (A1 style), isLocal = false
            string retrievedFormula = validation.GetFormula1(false, false);

            // Verify that the retrieved formula references cell A1
            bool referencesA1 = retrievedFormula != null &&
                               (retrievedFormula.Equals("A1", StringComparison.OrdinalIgnoreCase) ||
                                retrievedFormula.Equals("=A1", StringComparison.OrdinalIgnoreCase));

            Console.WriteLine($"Retrieved Formula: {retrievedFormula}");
            Console.WriteLine($"Does the formula reference A1? {(referencesA1 ? "Yes" : "No")}");

            // Save the workbook (optional, demonstrates lifecycle compliance)
            workbook.Save("ValidationReferenceCheck.xlsx");
        }
    }
}
