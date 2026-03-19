using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadWorkbookWithNonEnglishFormulaNotice
    {
        public static void Run()
        {
            // Create LoadOptions for XLSX file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Skip parsing formulas on load – useful for non‑English users where formulas may be in a different locale
            loadOptions.ParsingFormulaOnOpen = false;

            // Optionally set the culture that matches the source workbook (e.g., French)
            loadOptions.CultureInfo = new CultureInfo("fr-FR");

            // Load the workbook with the specified options
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Access the first worksheet to demonstrate that the workbook is loaded
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Output the raw formula string (parsing was skipped)
            Console.WriteLine("Cell A1 raw formula: " + cell.Formula);
            Console.WriteLine("Cell A1 value (may be empty until calculation): " + cell.Value);

            // If needed, calculate formulas later
            workbook.CalculateFormula();

            // Save the workbook after any modifications
            workbook.Save("output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadWorkbookWithNonEnglishFormulaNotice.Run();
        }
    }
}