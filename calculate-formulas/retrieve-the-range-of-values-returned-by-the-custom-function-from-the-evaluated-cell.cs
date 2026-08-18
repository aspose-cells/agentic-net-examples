// Title: Aspose.Cells for .NET – Retrieve a 2‑D array from a custom function and access its argument range
// Description: Demonstrates how to implement a custom calculation engine (MyArrayFunctionEngine) that returns a 2 × 2 object[,] array for the MYFUNC function, evaluate the formula with CalculationOptions, read the array via cell.Value, and use GetPrecedents and ReferredArea.GetValues to obtain the values of the passed range (A1:B2).
// Keywords: Aspose.Cells | C# | custom function | AbstractCalculationEngine | 2D array result | CalculateFormula | cell.Value | GetPrecedents | ReferredArea | Excel custom engine | matrix return
// Common Searches: Aspose.Cells custom function return array | how to get 2d array from custom function Aspose.Cells | GetPrecedents example C# | retrieve argument range values Aspose.Cells | custom calculation engine Aspose.Cells .NET
// Developer Intent: Extract the multi‑dimensional array produced by a custom Excel function and read the values of the range supplied as its argument.
// Use Cases: Create a custom calculation engine that returns a matrix and read the matrix after workbook calculation. | Identify the range passed to a custom function using GetPrecedents and extract its cell values with ReferredArea.GetValues. | Display both the function's result array and the source range data for debugging or reporting.
// AI Prompts: Generate C# code with Aspose.Cells that defines a custom function returning a 3 × 3 array and prints the array from the formula cell. | Show how to use ReferredAreaCollection and ReferredArea.GetValues to fetch the values of a range supplied to a custom function in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom engine that returns a 2‑D array as the result of MYFUNC
    // Demonstrates how to implement a custom calculation engine (MyArrayFunctionEngine) that returns a 2 × 2 object[,] array for the MYFUNC function, evaluate the formula with CalculationOptions, read the array via cell.Value, and use GetPrecedents and ReferredArea.GetValues to obtain the values of the passed range (A1:B2).
    class MyArrayFunctionEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Ensure we handle only the expected custom function
            if (data.FunctionName != null &&
                data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Example: return a 2x2 array with incremental numbers
                object[,] result = new object[2, 2];
                result[0, 0] = 1;
                result[0, 1] = 2;
                result[1, 0] = 3;
                result[1, 1] = 4;

                // Assign the array to the CalculatedValue property
                data.CalculatedValue = result;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some dummy data (not used by the custom function but required for a valid sheet)
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].PutValue(30);
            cells["B2"].PutValue(40);

            // Set a formula that calls the custom function
            // The function does not need parameters for this demo, but we include a range to illustrate GetPrecedents
            Cell formulaCell = cells["C1"];
            formulaCell.Formula = "=MYFUNC(A1:B2)";

            // ---------- Configure calculation options with the custom engine ----------
            CalculationOptions options = new CalculationOptions();
            options.CustomEngine = new MyArrayFunctionEngine();

            // Perform calculation – the custom engine will be invoked
            workbook.CalculateFormula(options);

            // ---------- Retrieve the range of values returned by the custom function ----------
            // Get the precedents of the formula cell; the first ReferredArea corresponds to the argument range (A1:B2)
            ReferredAreaCollection precedents = formulaCell.GetPrecedents();

            if (precedents != null && precedents.Count > 0)
            {
                // The custom function itself returns an array, not the argument range.
                // To obtain the array result, we use the cell's Value directly (it holds the array object).
                object result = formulaCell.Value;

                // The result can be a single value, a 1‑D array, or a 2‑D array.
                // In this example we expect a 2‑D array.
                if (result is object[,] multiArray)
                {
                    Console.WriteLine("Custom function returned a 2‑D array:");
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
                    Console.WriteLine("Custom function returned: " + result);
                }

                // Additionally, demonstrate retrieving values from the argument range using ReferredArea.GetValues()
                ReferredArea argArea = precedents[0];
                object argValues = argArea.GetValues(true); // calculate formulas inside the range if any

                if (argValues is object[,] argArray)
                {
                    Console.WriteLine("\nValues of the argument range (A1:B2):");
                    for (int r = 0; r < argArray.GetLength(0); r++)
                    {
                        for (int c = 0; c < argArray.GetLength(1); c++)
                        {
                            Console.Write(argArray[r, c] + "\t");
                        }
                        Console.WriteLine();
                    }
                }
            }
            else
            {
                Console.WriteLine("No precedents found for the formula cell.");
            }

            // ---------- Save the workbook (optional) ----------
            workbook.Save("CustomFunctionResult.xlsx");
        }
    }
}
