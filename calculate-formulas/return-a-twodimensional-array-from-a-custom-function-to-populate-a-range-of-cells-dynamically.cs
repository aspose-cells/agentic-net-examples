using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate source data (A1:B2) that the custom function will use
            cells["A1"].PutValue(1);
            cells["A2"].PutValue(2);
            cells["B1"].PutValue(3);
            cells["B2"].PutValue(4);

            // Define a formula that returns a two‑dimensional array.
            // Here we use TRANSPOSE as a stand‑in for a custom function that would
            // produce a 2‑D result based on the source range.
            string arrayFormula = "=TRANSPOSE(A1:B2)";

            // Calculate the array formula to obtain the resulting object[][].
            // This simulates a custom function returning a 2‑D array.
            CalculationOptions calcOptions = new CalculationOptions();
            object[][] resultArray = sheet.CalculateArrayFormula(arrayFormula, calcOptions);

            // Set the dynamic array formula in cell D1 and supply the pre‑calculated values.
            // calculateRange = false tells Aspose to use the dimensions of resultArray.
            // calculateValue = false prevents re‑calculation because we already have the values.
            Cell targetCell = cells["D1"];
            targetCell.SetDynamicArrayFormula(
                arrayFormula,
                new FormulaParseOptions(),
                resultArray,
                calculateRange: false,
                calculateValue: false);

            // Refresh dynamic array formulas so that the spill range is created.
            workbook.RefreshDynamicArrayFormulas(true);

            // Optional: verify the spilled values (D1:E2)
            Console.WriteLine("Spilled dynamic array values:");
            for (int r = 0; r < resultArray.Length; r++)
            {
                for (int c = 0; c < resultArray[r].Length; c++)
                {
                    Console.Write(cells[r, 3 + c].Value + "\t"); // D column index = 3
                }
                Console.WriteLine();
            }

            // Save the workbook
            workbook.Save("DynamicArrayResult.xlsx");
        }
    }
}