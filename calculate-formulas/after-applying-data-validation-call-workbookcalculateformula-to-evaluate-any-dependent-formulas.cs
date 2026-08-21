// Title: Aspose.Cells .NET: Add Data Validation with Cell Reference and Recalculate Formulas
// Description: Creates a new workbook, defines a whole‑number validation for A1 with a lower bound of 10 and an upper bound that references B1, assigns a formula to C1 that depends on A1, calls Workbook.CalculateFormula to update dependent cells, outputs the results, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | data validation | dynamic upper bound | cell reference validation | Workbook.CalculateFormula | formula recalculation | Excel automation | save workbook
// Common Searches: Aspose.Cells recalculate formulas after validation | set data validation upper bound to another cell Aspose.Cells | Workbook.CalculateFormula example C# | how to use cell reference in data validation Aspose.Cells | update dependent formulas in Aspose.Cells workbook
// Developer Intent: Apply a data‑validation rule that uses another cell as the upper limit, then recalculate any formulas that depend on the validated cell.
// Use Cases: Enforce user input in A1 to be between a constant and a value entered in B1. | Automatically update calculations (e.g., C1 = A1 × 2) after the validation rule is applied. | Generate an Excel file where the final values reflect the latest validation constraints without manual recalculation.
// AI Prompts: Show me C# code using Aspose.Cells to add a whole‑number validation to A1 with a lower bound of 10 and an upper bound that references B1, then recalculate dependent formulas. | How can I call Workbook.CalculateFormula after setting data validation so that a formula in C1 updates automatically? | Provide an Aspose.Cells example that validates a cell against another cell and ensures the workbook saves the evaluated results.

using System;
using Aspose.Cells;

namespace AsposeCellsDataValidationAndCalculate
{
    // Creates a new workbook, defines a whole‑number validation for A1 with a lower bound of 10 and an upper bound that references B1, assigns a formula to C1 that depends on A1, calls Workbook.CalculateFormula to update dependent cells, outputs the results, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 1. Add a sample value that will be used in validation
            // -------------------------------------------------
            cells["B1"].PutValue(50); // Reference value for validation

            // -------------------------------------------------
            // 2. Create a data validation rule for cell A1
            //    The rule requires the entered value to be between 10 and the value in B1
            // -------------------------------------------------
            int validationIndex = sheet.Validations.Add(new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 0,
                EndColumn = 0
            });
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "10";          // Lower bound
            validation.Formula2 = "=B1";        // Upper bound (reference to B1)

            // -------------------------------------------------
            // 3. Set a formula that depends on the validated cell (A1)
            //    For example, multiply the validated value by 2
            // -------------------------------------------------
            cells["C1"].Formula = "=A1*2";

            // -------------------------------------------------
            // 4. Calculate all formulas so that dependent cells are updated
            //    This satisfies the requirement: after applying data validation,
            //    call Workbook.CalculateFormula()
            // -------------------------------------------------
            workbook.CalculateFormula();

            // -------------------------------------------------
            // 5. Output the calculated result to the console
            // -------------------------------------------------
            Console.WriteLine("Value in A1 (if valid): " + cells["A1"].StringValue);
            Console.WriteLine("Calculated value in C1: " + cells["C1"].StringValue);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("DataValidationAndCalculate.xlsx", SaveFormat.Xlsx);
        }
    }
}
