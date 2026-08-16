// Title: Validate User‑Provided Excel Formula Syntax with Aspose.Cells in C#
// Description: Shows how to verify the syntax of an Excel formula string using Aspose.Cells before assigning it to a worksheet cell. The IsFormulaValid method calls Worksheet.CalculateFormula inside a try‑catch block; if no exception is thrown the formula is set, the workbook is calculated, and the file is saved.
// Keywords: Aspose.Cells | C# | .NET | Excel formula validation | formula syntax check | Worksheet.CalculateFormula | validate user formula | exception handling | Excel automation
// Common Searches: Aspose.Cells validate formula syntax C# | check Excel formula validity .NET | catch invalid formula exception Aspose.Cells | verify formula string before setting cell | how to test Excel formula with Aspose.Cells
// Developer Intent: The developer wants to ensure a user‑provided formula is syntactically correct before inserting it into a worksheet cell.
// Use Cases: Validate formulas entered through a UI to prevent runtime errors when saving Excel files. | Process large batches of dynamically generated formulas, skipping or logging those that fail validation. | Enable end‑users to create custom calculations while guaranteeing only valid formulas are stored.
// AI Prompts: Write a C# method that validates an Excel formula string using Aspose.Cells without modifying the worksheet, returning detailed error information. | Show how to catch and log the specific exception thrown by Worksheet.CalculateFormula when a formula is invalid. | Provide an alternative approach to syntax‑check a formula using Aspose.Cells' FormulaParser class.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaValidation
{
    // Shows how to verify the syntax of an Excel formula string using Aspose.Cells before assigning it to a worksheet cell. The IsFormulaValid method calls Worksheet.CalculateFormula inside a try‑catch block; if no exception is thrown the formula is set, the workbook is calculated, and the file is saved.
    class Program
    {
        // Validates a formula string by attempting to calculate it.
        // Returns true if the formula is syntactically correct, false otherwise.
        static bool IsFormulaValid(Worksheet sheet, string formula)
        {
            try
            {
                // Attempt to calculate the formula directly.
                // If the formula is invalid, an exception will be thrown.
                sheet.CalculateFormula(formula);
                return true;
            }
            catch (Exception)
            {
                // Invalid formula syntax.
                return false;
            }
        }

        static void Main(string[] args)
        {
            // Example user‑provided formula.
            string userFormula = "=SUM(A1:A5)";   // Change this to test different inputs.
            string targetCell = "B1";

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some data that the formula may reference.
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1:A5 = 1,2,3,4,5
            }

            // Validate the formula before inserting.
            if (IsFormulaValid(worksheet, userFormula))
            {
                // Formula is valid – set it to the target cell.
                cells[targetCell].Formula = userFormula;
                // Calculate all formulas in the workbook to obtain the result.
                workbook.CalculateFormula();
                Console.WriteLine($"Formula set successfully in {targetCell}. Result: {cells[targetCell].Value}");
            }
            else
            {
                Console.WriteLine("The provided formula is invalid and was not inserted.");
            }

            // Save the workbook (optional).
            workbook.Save("ValidatedFormula.xlsx");
        }
    }
}
