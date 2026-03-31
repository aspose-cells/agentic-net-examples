using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLoadingDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Create LoadOptions and configure formula‑related settings
            LoadOptions loadOptions = new LoadOptions();
            // Skip parsing formulas on load – formulas remain as raw strings
            loadOptions.ParsingFormulaOnOpen = false;
            // Preserve any padding spaces or line breaks inside formulas
            loadOptions.PreservePaddingSpacesInFormula = true;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Demonstrate that the formula was not parsed (value is default)
            Cell cell = sheet.Cells["A1"];
            Console.WriteLine("Cell A1 formula: " + cell.Formula);
            Console.WriteLine("Cell A1 value (before calculation): " + (cell.Value ?? "null"));

            // If needed, calculate formulas manually
            workbook.CalculateFormula();

            // Show the calculated value after explicit calculation
            Console.WriteLine("Cell A1 value (after calculation): " + (cell.Value ?? "null"));

            // Save the workbook to a new file (preserving the original formula formatting)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved to: " + outputPath);
        }
    }
}