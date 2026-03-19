using System;
using Aspose.Cells;

namespace AsposeCellsSharedFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string inputPath = "InputWorkbook.xlsx";
            // Path for the output workbook
            string outputPath = "OutputWorkbook.xlsx";

            // Configure load options to skip automatic formula parsing on open
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = false
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data in column A (if not already present)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1:A10 = 1..10
            }

            // Define a shared formula that doubles the value in column A
            string sharedFormula = "=A1*2";

            // Create formula parse options (default parsing)
            FormulaParseOptions parseOptions = new FormulaParseOptions
            {
                Parse = true,          // Ensure the formula is parsed
                R1C1Style = false,     // Use A1 style
                LocaleDependent = false
            };

            // Apply the shared formula starting at B1, covering 10 rows and 1 column
            Cell targetCell = cells["B1"];
            targetCell.SetSharedFormula(sharedFormula, 10, 1, parseOptions);

            // After setting the shared formula, parse any formulas that were not parsed earlier
            workbook.ParseFormulas(false);

            // Calculate all formulas to obtain results
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }
}