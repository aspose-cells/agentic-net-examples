using System;
using Aspose.Cells;

namespace AsposeCellsFormulaTextDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // LoadOptions allow control over how formulas are parsed on open.
            // Here we keep the default behavior (ParsingFormulaOnOpen = true) so that
            // any existing formulas are parsed immediately.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = true;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value into A1 – this will be the target of FORMULATEXT
            sheet.Cells["A1"].PutValue(123);

            // Set a formula in B1 that uses the Excel FUNCTION FORMULATEXT.
            // FORMULATEXT returns the formula string of the referenced cell.
            // Since A1 contains a constant, FORMULATEXT will return the constant itself.
            sheet.Cells["B1"].Formula = "=FORMULATEXT(A1)";

            // Calculate all formulas in the workbook. This will evaluate FORMULATEXT.
            workbook.CalculateFormula();

            // Retrieve and display the result of FORMULATEXT from B1.
            // The result is a string representation of the formula (or value) in A1.
            Console.WriteLine("Result of FORMULATEXT(A1) in B1: " + sheet.Cells["B1"].StringValue);

            // Optionally, save the workbook to see the formula and its result persisted.
            workbook.Save("output.xlsx");
        }
    }
}