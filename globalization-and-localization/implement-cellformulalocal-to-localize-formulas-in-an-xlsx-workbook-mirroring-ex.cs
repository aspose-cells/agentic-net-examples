using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Set the workbook's locale to German (Germany) to demonstrate localization
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and a target cell (A1)
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // -----------------------------------------------------------------
            // 1. Set a formula using the standard (English) notation.
            // -----------------------------------------------------------------
            cell.Formula = "=SUM(B1:C1)";

            // Display the formula in both standard and localized forms
            Console.WriteLine("After setting standard formula:");
            Console.WriteLine($"Standard Formula   : {cell.Formula}");
            Console.WriteLine($"Localized Formula  : {cell.FormulaLocal}");
            Console.WriteLine();

            // -----------------------------------------------------------------
            // 2. Set a formula using the localized (German) notation via FormulaLocal.
            // -----------------------------------------------------------------
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Display the formulas again to show the conversion
            Console.WriteLine("After setting localized formula:");
            Console.WriteLine($"Standard Formula   : {cell.Formula}");
            Console.WriteLine($"Localized Formula  : {cell.FormulaLocal}");
            Console.WriteLine();

            // -----------------------------------------------------------------
            // 3. Use GetFormula to retrieve formulas with explicit localization flags.
            // -----------------------------------------------------------------
            string englishFormula = cell.GetFormula(false, false); // not R1C1, not local
            string germanFormula = cell.GetFormula(false, true);   // not R1C1, local

            Console.WriteLine("Using GetFormula method:");
            Console.WriteLine($"English (standard) : {englishFormula}");
            Console.WriteLine($"German (localized) : {germanFormula}");
            Console.WriteLine();

            // -----------------------------------------------------------------
            // 4. Calculate the workbook to ensure the formula result is up‑to‑date.
            // -----------------------------------------------------------------
            workbook.CalculateFormula();

            // Show the calculated value
            Console.WriteLine($"Calculated Value in A1: {cell.Value}");
            Console.WriteLine();

            // -----------------------------------------------------------------
            // 5. Save the workbook (lifecycle: save)
            // -----------------------------------------------------------------
            string outputPath = "FormulaLocalDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}