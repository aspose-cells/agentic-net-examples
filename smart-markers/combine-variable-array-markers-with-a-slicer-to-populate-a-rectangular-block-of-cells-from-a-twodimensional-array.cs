// Title: C# – Fill a 3×3 range with a pre‑calculated 2‑D array using SetDynamicArrayFormula and a slicer in Aspose.Cells
// Description: Demonstrates how to create a workbook, define a 3 × 3 object array, place a slicer value, and apply SetDynamicArrayFormula with the SEQUENCE(3,3) formula to spill the array into A5:C7 without recalculating the range or values. The example also shows refreshing dynamic‑array formulas and saving the file.
// Keywords: Aspose.Cells C# | .NET dynamic array formula | SetDynamicArrayFormula | SEQUENCE spill range | 2D object array | variable array markers | slicer integration | populate rectangular block | pre‑calculated data matrix | RefreshDynamicArrayFormulas
// Common Searches: Aspose.Cells SetDynamicArrayFormula example C# | populate cells from 2D array Aspose.Cells | use slicer with dynamic array formula .NET | disable range calculation Aspose.Cells | dynamic array spill range C# Aspose
// Developer Intent: Insert a 3 × 3 block of values into a worksheet by applying a dynamic array formula that uses a pre‑computed 2‑D array and optionally references a slicer cell, while preventing automatic range and value recalculation.
// Use Cases: Load a pre‑processed data matrix into a spill range without triggering extra calculations. | Combine a slicer or dropdown cell value with a dynamic array to drive conditional data population. | Refresh dynamic‑array formulas after setting them to materialize the spill range in the saved workbook.
// AI Prompts: Generate C# code that uses Aspose.Cells to set a dynamic array formula for a 4 × 5 object array, disabling automatic range calculation. | Show how to link a slicer cell value to a SEQUENCE‑based dynamic array formula in Aspose.Cells and refresh the spill range. | Provide an example where SetDynamicArrayFormula is called with calculateRange:false and calculateValue:true to recalculate values while preserving dimensions.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, define a 3 × 3 object array, place a slicer value, and apply SetDynamicArrayFormula with the SEQUENCE(3,3) formula to spill the array into A5:C7 without recalculating the range or values. The example also shows refreshing dynamic‑array formulas and saving the file.
    class VariableArrayWithSlicerDemo
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Define a two‑dimensional array that we want to populate into the sheet
            //    The array is 3 rows × 3 columns
            object[][] data = new object[3][];
            data[0] = new object[] { 10, 20, 30 };
            data[1] = new object[] { 40, 50, 60 };
            data[2] = new object[] { 70, 80, 90 };

            // 3. Place a slicer value (for demonstration) in cell D1.
            //    In a real scenario this could be a dropdown or a slicer linked to a table.
            cells["D1"].PutValue("SliceValue");

            // 4. Set a dynamic array formula in cell A5.
            //    The formula itself is a simple SEQUENCE that creates a 3×3 spill range.
            //    We provide the pre‑calculated values (the 2‑D array) so that the cells are filled
            //    directly without re‑calculating the formula.
            Cell target = cells["A5"];
            string formula = "=SEQUENCE(3,3)";                     // creates a 3‑row, 3‑column spill
            FormulaParseOptions parseOptions = new FormulaParseOptions(); // default options

            // calculateRange = false  -> use the dimensions of the supplied 'data' array
            // calculateValue = false  -> do not recalculate, use the supplied values
            target.SetDynamicArrayFormula(formula, parseOptions, data, calculateRange: false, calculateValue: false);

            // 5. Refresh dynamic array formulas so that the spill range is materialised.
            //    The 'true' flag also calculates the values (already supplied) for completeness.
            workbook.RefreshDynamicArrayFormulas(true);

            // 6. Optionally, calculate the whole workbook (not strictly required here)
            workbook.CalculateFormula();

            // 7. Save the workbook
            workbook.Save("VariableArrayWithSlicerDemo.xlsx");
        }
    }
}
