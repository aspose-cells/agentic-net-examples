using System;
using Aspose.Cells;

namespace AsposeCellsLoadOptionsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Create LoadOptions instance
            LoadOptions loadOptions = new LoadOptions();

            // Do not parse formulas on open – this keeps the original formula strings
            // and prevents calculation of their results, while still loading all formatting.
            loadOptions.ParsingFormulaOnOpen = false;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet to demonstrate that formatting is preserved
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Output the cell's raw value (formula string) and its style name
            Console.WriteLine($"Cell A1 raw value: {cell.Value}");
            Console.WriteLine($"Cell A1 style name: {cell.GetStyle().Name}");

            // Save the workbook to a new file – formatting remains intact,
            // and formula results are not calculated.
            string outputPath = "output_preserve_format.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}