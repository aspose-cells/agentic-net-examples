using System;
using Aspose.Cells;

namespace AsposeCellsFormulaCalculation
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Configure load options to skip formula parsing during file opening.
            // This improves load performance when the workbook contains many formulas.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = false;

            // Load the workbook with the specified options.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Parse all formulas that were not parsed during the load operation.
            // The 'ignoreError' flag is set to false so that any invalid formula will raise an exception.
            workbook.ParseFormulas(false);

            // Calculate all formulas in the workbook.
            workbook.CalculateFormula();

            // (Optional) Save the workbook after calculation.
            // The Save method is part of the standard lifecycle and complies with the required rules.
            workbook.Save("output.xlsx", SaveFormat.Xlsx);

            // Demonstrate that formulas have been evaluated by printing a sample cell value.
            // Adjust the cell reference as needed for your specific workbook.
            Console.WriteLine("Calculated value of A1: " + workbook.Worksheets[0].Cells["A1"].Value);
        }
    }
}