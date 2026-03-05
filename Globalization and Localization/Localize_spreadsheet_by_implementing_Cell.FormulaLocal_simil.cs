using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file (must exist)
            string inputPath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook locale to German (de-DE) for demonstration
            // This influences how FormulaLocal is interpreted and displayed
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and a target cell (A1)
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Set a formula using the standard (English) syntax
            cell.Formula = "=SUM(B1:C1)";

            // Display the formula in both standard and localized forms
            Console.WriteLine("After setting standard formula:");
            Console.WriteLine("Standard Formula : " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Now set the formula using the German localized syntax
            // In German, the SUM function is "SUMME"
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Display the formulas again to show the conversion
            Console.WriteLine("\nAfter setting localized formula:");
            Console.WriteLine("Standard Formula : " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Optionally calculate the workbook so that the cell shows the result
            workbook.CalculateFormula();

            // Show the calculated value
            Console.WriteLine("\nCalculated Value in A1: " + cell.Value);

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
        }
    }
}