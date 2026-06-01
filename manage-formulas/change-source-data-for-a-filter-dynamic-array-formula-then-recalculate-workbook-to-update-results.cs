using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayRefreshDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate source data for the FILTER formula
            //    A column: values to filter, B column: criteria
            cells["A2"].PutValue(10);
            cells["A3"].PutValue(20);
            cells["A4"].PutValue(30);
            cells["A5"].PutValue(40);

            cells["B2"].PutValue(5);
            cells["B3"].PutValue(15);
            cells["B4"].PutValue(25);
            cells["B5"].PutValue(35);

            // 3. Set a FILTER dynamic array formula in C2.
            //    It will return values from A2:A5 where B2:B5 > 20.
            //    Use SetDynamicArrayFormula (rule) and enable calculation.
            Cell formulaCell = cells["C2"];
            string filterFormula = "=FILTER(A2:A5, B2:B5>20)";
            formulaCell.SetDynamicArrayFormula(filterFormula, new FormulaParseOptions(), true);

            // 4. Calculate the workbook so the initial spill range is populated.
            workbook.CalculateFormula();

            // 5. Output initial results
            Console.WriteLine("Initial FILTER results:");
            for (int row = 2; row <= 6; row++) // spill may occupy rows 2..6
            {
                Cell c = cells.CheckCell(row, 2); // column C index = 2
                if (c != null && c.IsFormula) break; // stop at first empty cell
                Console.WriteLine($"C{row}: {c?.Value ?? "Empty"}");
            }

            // 6. Change source data that influences the FILTER formula.
            //    Update B4 from 25 to 10, so now only B5 > 20.
            cells["B4"].PutValue(10);

            // 7. Refresh dynamic array formulas and recalculate affected cells.
            //    Pass true to also calculate the spilled range.
            workbook.RefreshDynamicArrayFormulas(true);

            // 8. Output refreshed results
            Console.WriteLine("\nAfter source data change and refresh:");
            for (int row = 2; row <= 6; row++)
            {
                Cell c = cells.CheckCell(row, 2);
                if (c != null && c.IsFormula) break;
                Console.WriteLine($"C{row}: {c?.Value ?? "Empty"}");
            }

            // 9. Save the workbook (lifecycle save)
            workbook.Save("DynamicArrayRefreshResult.xlsx");
        }
    }
}