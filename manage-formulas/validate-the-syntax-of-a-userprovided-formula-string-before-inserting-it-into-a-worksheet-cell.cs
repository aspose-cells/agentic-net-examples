using System;
using Aspose.Cells;

namespace AsposeCellsFormulaValidation
{
    class Program
    {
        // Validates the syntax of a formula string.
        // Returns true if the formula can be parsed without error.
        static bool ValidateFormula(string formula)
        {
            // Create a temporary workbook for validation only.
            Workbook tempWorkbook = new Workbook();
            Worksheet tempSheet = tempWorkbook.Worksheets[0];
            Cells tempCells = tempSheet.Cells;

            // Set the formula with parsing disabled.
            // The third parameter (value) is set to null because we only care about parsing.
            tempCells["A1"].SetFormula(formula, new FormulaParseOptions() { Parse = false }, null);

            try
            {
                // Attempt to parse all formulas in the workbook.
                // ignoreError = false will cause an exception if the formula is invalid.
                tempWorkbook.ParseFormulas(false);
                return true; // No exception -> formula is syntactically correct.
            }
            catch (Exception)
            {
                return false; // Parsing failed -> invalid formula.
            }
        }

        static void Main(string[] args)
        {
            // Example user‑provided formula.
            string userFormula = "=SUM(A1:A5)";

            // Validate the formula before inserting.
            if (ValidateFormula(userFormula))
            {
                // Create the actual workbook where the formula will be stored.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Insert the validated formula into cell B2.
                cells["B2"].SetFormula(userFormula, new FormulaParseOptions() { Parse = true }, null);

                // Optionally calculate to obtain the result.
                workbook.CalculateFormula();

                Console.WriteLine("Formula inserted successfully. Result in B2: " + cells["B2"].Value);
                // Save the workbook.
                workbook.Save("ValidatedFormula.xlsx");
            }
            else
            {
                Console.WriteLine("The provided formula is syntactically invalid and was not inserted.");
            }
        }
    }
}