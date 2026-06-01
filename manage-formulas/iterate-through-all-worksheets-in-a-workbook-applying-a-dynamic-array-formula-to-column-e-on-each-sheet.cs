using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // replace with new Workbook("input.xlsx") to load

            // Example: add a second worksheet to demonstrate iteration over multiple sheets
            workbook.Worksheets.Add("Sheet2");

            // Define the dynamic array formula to be applied.
            // This example uses SEQUENCE to generate a vertical array of numbers.
            // Adjust the formula as needed for your scenario.
            string dynamicArrayFormula = "=SEQUENCE(10)";

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Apply the dynamic array formula to the first cell of column E (E1)
                // The formula will spill into the cells below as needed.
                sheet.Cells["E1"].SetDynamicArrayFormula(
                    dynamicArrayFormula,
                    new FormulaParseOptions(), // default parse options
                    true);                     // calculate the spilled values immediately
            }

            // Refresh dynamic array formulas to ensure spill ranges are up‑to‑date
            workbook.RefreshDynamicArrayFormulas(true);

            // Optionally calculate all other formulas in the workbook
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save("DynamicArrayResult.xlsx");
        }
    }
}