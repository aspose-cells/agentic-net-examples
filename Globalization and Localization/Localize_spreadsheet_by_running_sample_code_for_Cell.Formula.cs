using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook (replace with actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook's default locale to German (Germany)
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and cell A1
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Set a formula using the standard (English) syntax
            cell.Formula = "=SUM(B1:C1)";

            // Display the formula in both standard and localized forms
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula (FormulaLocal): " + cell.FormulaLocal);

            // Set the formula using the German localized syntax
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Display the formulas again to show the difference
            Console.WriteLine("\nAfter setting FormulaLocal:");
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula (FormulaLocal): " + cell.FormulaLocal);

            // Demonstrate GetFormula with localization flag
            Console.WriteLine("\nUsing GetFormula:");
            Console.WriteLine("English formula (isLocal = false): " + cell.GetFormula(false, false));
            Console.WriteLine("Localized formula (isLocal = true): " + cell.GetFormula(false, true));

            // Save the modified workbook (replace with desired output path)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}