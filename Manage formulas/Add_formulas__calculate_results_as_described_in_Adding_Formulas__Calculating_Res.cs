using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file that contains formulas
            string sourcePath = "input.xlsx";

            // Configure load options to skip formula parsing during the initial load.
            // This speeds up loading large workbooks when we intend to parse and calculate later.
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = false
            };

            // Load the workbook with the specified options.
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // At this point formulas are stored as raw strings and not yet parsed.
            // Parse all formulas now. The 'ignoreError' flag is set to false so any invalid
            // formula will raise an exception, helping to catch issues early.
            workbook.ParseFormulas(false);

            // After parsing, calculate all formulas in the workbook.
            workbook.CalculateFormula();

            // Optionally, you can access calculated values to verify.
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("Calculated value of A1: " + sheet.Cells["A1"].Value);
            Console.WriteLine("Calculated value of B1: " + sheet.Cells["B1"].Value);

            // Save the workbook with calculated results.
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}