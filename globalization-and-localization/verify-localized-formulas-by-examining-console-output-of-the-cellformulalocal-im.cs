using System;
using Aspose.Cells;

namespace VerifyLocalizedFormulas
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the workbook locale to German (Germany)
            workbook.Settings.Region = CountryCode.Germany;

            // Access cell A1
            Cell cell = worksheet.Cells["A1"];

            // Set formula using the standard (English) notation
            cell.Formula = "=SUM(B1:C1)";

            // Output the standard and localized formulas
            Console.WriteLine("After setting standard formula:");
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Set formula using the German localized notation
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Output the formulas again to see the conversion
            Console.WriteLine("\nAfter setting localized formula:");
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Demonstrate GetFormula with localization flags
            Console.WriteLine("\nUsing GetFormula:");
            Console.WriteLine("English formula (isLocal=false): " + cell.GetFormula(false, false));
            Console.WriteLine("Localized formula (isLocal=true): " + cell.GetFormula(false, true));

            // Save the workbook (optional, just to verify no errors on save)
            workbook.Save("LocalizedFormulaDemo.xlsx");
        }
    }
}