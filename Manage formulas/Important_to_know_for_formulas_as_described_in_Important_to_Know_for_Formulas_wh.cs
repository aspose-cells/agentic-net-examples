using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSX file that contains formulas
            string filePath = "input.xlsx";

            // Create LoadOptions and disable formula parsing on open.
            // This speeds up loading when you do not need immediate formula evaluation.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = false; // Skip parsing formulas during load

            // Load the workbook with the specified options using the supported constructor.
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Access the first worksheet and a cell that contains a formula.
            Worksheet sheet = workbook.Worksheets[0];
            Cell formulaCell = sheet.Cells["A1"];

            // At this point the formula text is available but the value is not calculated.
            Console.WriteLine("Cell A1 formula (as loaded): " + formulaCell.Formula);
            Console.WriteLine("Cell A1 value before parsing: " + (formulaCell.Value ?? "null"));

            // Parse all formulas that were not parsed during load.
            // The boolean parameter indicates whether to ignore errors in invalid formulas.
            workbook.ParseFormulas(ignoreError: true);

            // After parsing, the cell value is calculated.
            Console.WriteLine("Cell A1 value after parsing: " + (formulaCell.Value ?? "null"));

            // Optionally, recalculate all formulas in the workbook (e.g., if dependent cells changed).
            workbook.CalculateFormula();

            // Save the workbook to verify that formulas are now stored with calculated values.
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}