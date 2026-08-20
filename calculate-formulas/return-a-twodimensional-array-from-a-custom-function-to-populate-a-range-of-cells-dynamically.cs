// Title: Spill a 2‑D array from a custom function using Aspose.Cells SetDynamicArrayFormula in C#
// AI Prompts: Generate C# code that creates a Workbook, builds a two‑dimensional object array, and applies SetDynamicArrayFormula so the values spill from A1 without recalculation. | Show how to call workbook.RefreshDynamicArrayFormulas after setting a dynamic array formula and then save the file.
// Common Searches: Aspose.Cells SetDynamicArrayFormula example C# | C# spill multi‑cell array to Excel with Aspose.Cells | disable recalculation when using SetDynamicArrayFormula | RefreshDynamicArrayFormulas Aspose.Cells usage | custom function returning array in Aspose.Cells
// Tags: Aspose.Cells | SetDynamicArrayFormula | C# | dynamic array spill | RefreshDynamicArrayFormulas | Excel custom function

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayDemo
{
    // Demonstrates creating a workbook, preparing a 3 × 2 object array, using SetDynamicArrayFormula to spill the values starting at A1 without triggering recalculation, refreshing the dynamic array formulas, and saving the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Prepare a two‑dimensional array that we want to spill into the sheet
            // This mimics the result that a custom function would return
            object[][] arrayValues = new object[3][];
            arrayValues[0] = new object[] { 10, 20 };
            arrayValues[1] = new object[] { 30, 40 };
            arrayValues[2] = new object[] { 50, 60 };

            // Set a dynamic array formula in cell A1.
            // The formula name "MYARRAYFUNC()" is just a placeholder; we supply the pre‑calculated values.
            Cell targetCell = cells["A1"];
            targetCell.SetDynamicArrayFormula(
                "=MYARRAYFUNC()",                 // formula text (could be any valid Excel formula)
                new FormulaParseOptions(),        // parse options (default)
                arrayValues,                      // pre‑calculated 2‑D values to spill
                calculateRange: false,            // use the dimensions of arrayValues for the spill range
                calculateValue: false);           // do not recalculate; use supplied values

            // Refresh dynamic array formulas so that the spill range is materialized in the worksheet
            workbook.RefreshDynamicArrayFormulas(true);

            // Optional: verify the spilled values by printing them to the console
            Console.WriteLine("Spilled dynamic array values:");
            for (int r = 0; r < arrayValues.Length; r++)
            {
                for (int c = 0; c < arrayValues[r].Length; c++)
                {
                    Console.Write(cells[r, c].Value + "\t");
                }
                Console.WriteLine();
            }

            // Save the workbook to a file
            workbook.Save("DynamicArrayResult.xlsx");
        }
    }
}
