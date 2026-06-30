using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example – load workbook without parsing/calculating formulas
    class LoadWorkbookWithoutFormulaCalculation
    {
        static void Main()
        {
            // Path to the source workbook (replace with your actual file path)
            string filePath = "input.xlsx";

            // Create LoadOptions and disable formula parsing on open
            LoadOptions loadOptions = new LoadOptions
            {
                // When false, formulas are loaded as raw strings and not evaluated
                ParsingFormulaOnOpen = false
            };

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Demonstrate that the formula is preserved but not calculated
            Worksheet sheet = workbook.Worksheets[0];
            Cell formulaCell = sheet.Cells["A1"];

            Console.WriteLine("Cell A1 formula: " + formulaCell.Formula);
            Console.WriteLine("Cell A1 value (not calculated): " + formulaCell.Value);
        }
    }
}