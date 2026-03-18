using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set the workbook's default locale to German (de-DE)
            // This influences how locale‑dependent formulas are interpreted.
            workbook.Settings.Region = CountryCode.Germany;

            // ------------------------------------------------------------
            // Demonstrate standard (English) formula and its localized form
            // ------------------------------------------------------------
            Cell cellA1 = cells["A1"];
            // Set formula using the standard (English) syntax
            cellA1.Formula = "=SUM(B1:C1)";

            // Display both representations
            Console.WriteLine("Standard Formula (English): " + cellA1.Formula);
            Console.WriteLine("Localized Formula (German): " + cellA1.FormulaLocal);

            // Now set the formula using the German localized function name
            Cell cellA2 = cells["A2"];
            cellA2.FormulaLocal = "=SUMME(B2:C2)";

            // Verify that the internal (English) formula is automatically converted
            Console.WriteLine("\nAfter setting FormulaLocal:");
            Console.WriteLine("Standard Formula (English): " + cellA2.Formula);
            Console.WriteLine("Localized Formula (German): " + cellA2.FormulaLocal);

            // ------------------------------------------------------------
            // Apply locale‑specific formulas to a range of cells
            // ------------------------------------------------------------
            // Populate some data that the formulas will reference
            for (int i = 1; i <= 5; i++)
            {
                cells[$"B{i}"].PutValue(i * 10); // B1..B5 = 10,20,30,40,50
                cells[$"C{i}"].PutValue(i * 5);  // C1..C5 = 5,10,15,20,25
            }

            // Use FormulaLocal to set German SUMME formulas in column D
            for (int row = 1; row <= 5; row++)
            {
                Cell target = cells[$"D{row}"];
                // German syntax: =SUMME(Bx:Cx)
                target.FormulaLocal = $"=SUMME(B{row}:C{row})";
            }

            // Calculate all formulas so that results are stored
            workbook.CalculateFormula();

            // Output the calculated results for verification
            Console.WriteLine("\nCalculated results (German SUMME):");
            for (int row = 1; row <= 5; row++)
            {
                Console.WriteLine($"D{row} = {cells[$"D{row}"].Value}");
            }

            // ------------------------------------------------------------
            // Save the workbook to an XLSX file (lifecycle rule)
            // ------------------------------------------------------------
            workbook.Save("FormulaLocalDemo.xlsx");
            Console.WriteLine("\nWorkbook saved as 'FormulaLocalDemo.xlsx'.");
        }
    }
}