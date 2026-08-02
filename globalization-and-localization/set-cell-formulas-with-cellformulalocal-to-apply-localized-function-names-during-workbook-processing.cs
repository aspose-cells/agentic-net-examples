using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set the workbook region to German (locale where SUM is "SUMME")
            workbook.Settings.Region = CountryCode.Germany;

            // Put some sample values that will be used by the formula
            cells["B1"].PutValue(5);
            cells["C1"].PutValue(7);

            // Access cell A1
            Cell target = cells["A1"];

            // Set the formula using the standard (English) function name
            target.Formula = "=SUM(B1:C1)";

            // Display the formula in standard and localized forms before calculation
            Console.WriteLine("Standard Formula : " + target.Formula);
            Console.WriteLine("Localized Formula : " + target.FormulaLocal);

            // Now set the formula using the localized function name via FormulaLocal
            // In German the SUM function is called "SUMME"
            target.FormulaLocal = "=SUMME(B1:C1)";

            // Verify that both properties reflect the same underlying formula
            Console.WriteLine("\nAfter setting FormulaLocal:");
            Console.WriteLine("Standard Formula : " + target.Formula);
            Console.WriteLine("Localized Formula : " + target.FormulaLocal);

            // Calculate the workbook to evaluate the formula
            workbook.CalculateFormula();

            // Output the calculated result
            Console.WriteLine("\nCalculated Value in A1: " + target.Value);

            // Save the workbook (optional, demonstrates lifecycle rule usage)
            workbook.Save("FormulaLocalDemo.xlsx");
        }
    }
}