// Title: Aspose.Cells C# – Retrieve a 2‑D array from a formula's referenced range using ReferredArea.GetValues
// Description: Creates a workbook, fills A1:B2 with numbers, adds a SUM formula in C1, calculates the sheet, obtains the formula's precedents, selects the first ReferredArea (A1:B2) and calls GetValues to return the cell contents as an object[,] (2‑dimensional array). The array is printed and the workbook is saved.
// Keywords: Aspose.Cells ReferredArea GetValues | C# get 2D array from range | formula precedents Aspose.Cells | extract multiple cell values | object[,] Aspose.Cells example | calculate formulas C# | read precedent range values
// Common Searches: How to use ReferredArea.GetValues in Aspose.Cells C# | Get values of a referenced range as object[,] | Aspose.Cells retrieve precedent cells | C# extract 2‑dimensional array from formula range | Aspose.Cells GetValues multi‑cell example
// Developer Intent: Extract the values of a formula's referenced range as a two‑dimensional array with ReferredArea.GetValues.
// Use Cases: Read all cells of a summed range for custom post‑calculation logic. | Iterate over precedent values to validate or transform data before export. | Handle single‑cell and multi‑cell references uniformly when processing formula dependencies.
// AI Prompts: Write C# code that uses Aspose.Cells to obtain a 2‑D object[,] from the range referenced by a formula cell. | Show how to detect whether ReferredArea.GetValues returns a scalar or an object[,] and process each scenario. | Provide a complete example that calculates formulas, fetches precedents, and prints the extracted array of values.

using System;
using Aspose.Cells;

namespace ReferredAreaGetValuesDemo
{
    // Creates a workbook, fills A1:B2 with numbers, adds a SUM formula in C1, calculates the sheet, obtains the formula's precedents, selects the first ReferredArea (A1:B2) and calls GetValues to return the cell contents as an object[,] (2‑dimensional array). The array is printed and the workbook is saved.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a 2x2 range with sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].PutValue(30);
            sheet.Cells["B2"].PutValue(40);

            // Place a formula that references the range A1:B2
            Cell formulaCell = sheet.Cells["C1"];
            formulaCell.Formula = "=SUM(A1:B2)";

            // Ensure formulas are calculated (optional for GetValues)
            workbook.CalculateFormula();

            // Retrieve the referred areas (precedents) of the formula cell
            ReferredAreaCollection precedents = formulaCell.GetPrecedents();

            if (precedents != null && precedents.Count > 0)
            {
                // Get the first referred area (the range A1:B2)
                ReferredArea area = precedents[0];

                // Use GetValues() to obtain a 2‑dimensional array of the cell values
                object values = area.GetValues();

                // The returned object is a 2‑D array when the area covers multiple cells
                if (values is object[,] multiArray)
                {
                    Console.WriteLine("Values retrieved from the referred area:");
                    for (int r = 0; r < multiArray.GetLength(0); r++)
                    {
                        for (int c = 0; c < multiArray.GetLength(1); c++)
                        {
                            Console.Write(multiArray[r, c] + "\t");
                        }
                        Console.WriteLine();
                    }
                }
                else
                {
                    // Single‑cell area returns the cell value directly
                    Console.WriteLine("Single value: " + values);
                }
            }

            // Save the workbook (demonstrates the required save step)
            workbook.Save("ReferredAreaGetValuesDemo.xlsx");
        }
    }
}
