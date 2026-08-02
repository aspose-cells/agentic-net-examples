// Title: Validate an Excel formula string before inserting it into a cell with Aspose.Cells for .NET
// Description: Demonstrates how to use a temporary Workbook and the SetFormula Parse option to check a user‑provided formula for syntax errors. The method disables parsing, calls ParseFormulas, catches any exception, and only writes the formula to the target workbook when it is valid, then calculates and saves the file.
// Keywords: Aspose.Cells formula validation | C# Excel formula syntax check | SetFormula Parse false | ParseFormulas exception handling | validate user formula Aspose | Excel formula parsing .NET | temporary workbook validation
// Common Searches: how to validate an Excel formula with Aspose.Cells C# | check formula syntax before inserting cell Aspose | Aspose.Cells parse formula without error | C# validate user entered formula Aspose | prevent invalid formulas in Aspose.Cells workbook
// Developer Intent: Confirm that a formula supplied by a user or external source is syntactically correct before writing it to a worksheet cell.
// Use Cases: Web applications that accept formulas from end‑users and need to avoid runtime calculation errors. | Import pipelines that read formulas from CSV or JSON files and must ensure each formula is valid before applying it. | Batch processing of configuration‑driven formulas where only syntactically correct entries are written to the workbook.
// AI Prompts: Generate a C# method that validates an Excel formula string using Aspose.Cells without affecting the main workbook. | Show how to capture specific parsing exceptions from ParseFormulas and return a detailed error message. | Provide sample code that iterates over a list of formulas, validates each one, and inserts only the valid formulas into separate cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaValidation
{
    // Demonstrates how to use a temporary Workbook and the SetFormula Parse option to check a user‑provided formula for syntax errors. The method disables parsing, calls ParseFormulas, catches any exception, and only writes the formula to the target workbook when it is valid, then calculates and saves the file.
    class Program
    {
        // Validates the syntax of a formula string.
        // Returns true if the formula can be parsed without errors.
        static bool IsFormulaValid(string formula)
        {
            // Create a temporary workbook for validation.
            Workbook tempWorkbook = new Workbook();
            Worksheet tempSheet = tempWorkbook.Worksheets[0];
            Cell tempCell = tempSheet.Cells["A1"];

            // Set the formula with parsing disabled.
            // This prevents immediate parsing errors and allows batch parsing later.
            tempCell.SetFormula(formula, new FormulaParseOptions() { Parse = false }, null);

            try
            {
                // Parse all formulas in the workbook.
                // ignoreError = false => an exception is thrown for invalid formulas.
                tempWorkbook.ParseFormulas(false);
                return true; // No exception means the formula is syntactically correct.
            }
            catch (Exception)
            {
                return false; // Parsing failed, formula is invalid.
            }
        }

        static void Main(string[] args)
        {
            // Example user‑provided formula.
            string userFormula = "=SUM(A1:A5)";

            // Validate the formula before inserting.
            if (IsFormulaValid(userFormula))
            {
                // Create the actual workbook where the formula will be stored.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cell targetCell = sheet.Cells["B1"];

                // Insert the validated formula.
                targetCell.SetFormula(userFormula, new FormulaParseOptions() { Parse = true }, null);

                // Calculate to ensure the result is available.
                workbook.CalculateFormula();

                Console.WriteLine($"Formula inserted successfully. Result: {targetCell.Value}");
                // Save the workbook if needed.
                workbook.Save("ValidatedFormula.xlsx");
            }
            else
            {
                Console.WriteLine("The provided formula is invalid and was not inserted.");
            }
        }
    }
}
