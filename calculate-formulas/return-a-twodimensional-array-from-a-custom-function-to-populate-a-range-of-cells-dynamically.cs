// Title: C# – Return a 2‑D object array from a custom function and fill a dynamic‑array spill range with Aspose.Cells
// Description: This example shows how to create a workbook, use CalculateArrayFormula to evaluate a SEQUENCE formula into an object[][], and then assign the pre‑computed values to cell A1 with SetDynamicArrayFormula. After calling RefreshDynamicArrayFormulas the spill range is generated, printed, and the file is saved as DynamicArrayResult.xlsx.
// Keywords: Aspose.Cells C# dynamic array | CalculateArrayFormula object[][] | SetDynamicArrayFormula example | RefreshDynamicArrayFormulas | SEQUENCE function Aspose.Cells | populate spill range programmatically | pre‑calculate Excel array formula
// Common Searches: How to get a 2D object array from a formula in Aspose.Cells | SetDynamicArrayFormula without immediate calculation | RefreshDynamicArrayFormulas to create spill area | Populate Excel range with SEQUENCE result using Aspose.Cells | Aspose.Cells custom function returning object[][]
// Developer Intent: Generate a two‑dimensional object array from a formula and apply it to a dynamic‑array spill range in one operation.
// Use Cases: Pre‑compute the result of SEQUENCE (or any custom array formula) and inject it into a worksheet to avoid runtime recalculation. | Create a reusable method that returns object[][] and use SetDynamicArrayFormula to fill large tables efficiently. | Programmatically refresh dynamic‑array formulas after setting values so the spill range is correctly created and can be read back.
// AI Prompts: Write C# code that defines a custom function returning object[][] and uses SetDynamicArrayFormula to populate a range in Aspose.Cells. | Explain how to configure CalculationOptions to prevent NullReferenceException when calling CalculateArrayFormula and SetDynamicArrayFormula. | Show how to verify the spilled range and extract its values after setting a dynamic‑array formula with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayDemo
{
    // This example shows how to create a workbook, use CalculateArrayFormula to evaluate a SEQUENCE formula into an object[][], and then assign the pre‑computed values to cell A1 with SetDynamicArrayFormula. After calling RefreshDynamicArrayFormulas the spill range is generated, printed, and the file is saved as DynamicArrayResult.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Prepare input data for the dynamic‑array formula
                //    B1 will hold the number of rows for SEQUENCE
                cells["B1"].PutValue(3); // SEQUENCE will generate 3 rows

                // 3. Define the dynamic‑array formula (2 columns)
                string dynamicFormula = "=SEQUENCE(B1,2)";

                // 4. Calculate the array formula to obtain the 2‑D result
                //    CalculateArrayFormula returns object[][] (rows × columns)
                object[][] resultArray = sheet.CalculateArrayFormula(dynamicFormula, new CalculationOptions());

                // 5. Set the dynamic‑array formula in A1 and supply the pre‑calculated values
                //    Pass CalculationOptions to avoid internal NullReferenceException
                Cell targetCell = cells["A1"];
                targetCell.SetDynamicArrayFormula(
                    dynamicFormula,
                    new FormulaParseOptions(),
                    resultArray,
                    calculateRange: false,
                    calculateValue: false,
                    new CalculationOptions());

                // 6. Refresh dynamic‑array formulas so the spill range is created
                workbook.RefreshDynamicArrayFormulas(true);

                // 7. (Optional) Verify the spilled values by printing them
                CellArea spillArea = targetCell.GetArrayRange();
                Console.WriteLine("Spilled range: {0}:{1}",
                    CellsHelper.CellIndexToName(spillArea.StartRow, spillArea.StartColumn),
                    CellsHelper.CellIndexToName(spillArea.EndRow, spillArea.EndColumn));

                for (int r = spillArea.StartRow; r <= spillArea.EndRow; r++)
                {
                    for (int c = spillArea.StartColumn; c <= spillArea.EndColumn; c++)
                    {
                        Console.Write(cells[r, c].Value + "\t");
                    }
                    Console.WriteLine();
                }

                // 8. Save the workbook
                workbook.Save("DynamicArrayResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
