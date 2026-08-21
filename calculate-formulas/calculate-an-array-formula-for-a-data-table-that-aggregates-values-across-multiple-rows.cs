// Title: C# – Calculate an array formula with Aspose.Cells for .NET and retrieve the result
// Description: This example creates a workbook, fills cells A1:B4 with numbers, and uses Worksheet.CalculateArrayFormula to evaluate the array formula "=SUM(A1:B4)". The scalar sum is returned as a 1×1 object array. It also shows the overload that limits rows and columns by generating a 5‑row × 2‑column SEQUENCE array, iterating the result, and finally saving the workbook.
// Keywords: Aspose.Cells CalculateArrayFormula | C# array formula sum | SUM function Aspose.Cells | SEQUENCE function C# | retrieve array formula result | object[][] result handling | limit rows columns CalculateArrayFormula | Aspose.Cells .NET example | aggregate values across rows
// Common Searches: Aspose.Cells calculate array formula C# | Worksheet.CalculateArrayFormula example | How to get SUM result from array formula in Aspose.Cells | SEQUENCE function with CalculateArrayFormula overload | Retrieve scalar value from CalculateArrayFormula
// Developer Intent: Programmatically evaluate an array formula on a range and access the returned value(s) using Aspose.Cells for .NET.
// Use Cases: Sum all numbers in a data table (A1:B4) with an array formula and read the single scalar from the returned object[][] array. | Generate a multi‑cell array using the SEQUENCE function, limit the returned size to a specific number of rows and columns, and iterate the result for display or further processing. | Persist the workbook after performing array calculations to keep generated data.
// AI Prompts: Show C# code that calls Worksheet.CalculateArrayFormula to compute =SUM(A1:B4) and extracts the scalar sum from the object[][] result. | Provide an example of using CalculateArrayFormula with row and column limits to retrieve a 5×2 SEQUENCE array and print each element. | Explain how to handle the object[][] output when the array formula returns a single value versus a multi‑cell array in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsArrayFormulaDemo
{
    // This example creates a workbook, fills cells A1:B4 with numbers, and uses Worksheet.CalculateArrayFormula to evaluate the array formula "=SUM(A1:B4)". The scalar sum is returned as a 1×1 object array. It also shows the overload that limits rows and columns by generating a 5‑row × 2‑column SEQUENCE array, iterating the result, and finally saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (A1:B4)
            // Row 1
            cells["A1"].PutValue(1);
            cells["B1"].PutValue(2);
            // Row 2
            cells["A2"].PutValue(3);
            cells["B2"].PutValue(4);
            // Row 3
            cells["A3"].PutValue(5);
            cells["B3"].PutValue(6);
            // Row 4
            cells["A4"].PutValue(7);
            cells["B4"].PutValue(8);

            // Define an array formula that aggregates values across the whole table.
            // Here we use SUM to add all numbers in the range A1:B4.
            string arrayFormula = "=SUM(A1:B4)";

            // Create calculation options (default options are sufficient for this demo)
            CalculationOptions calcOptions = new CalculationOptions();

            // Calculate the array formula. The result is returned as a two‑dimensional object array.
            // Since the formula returns a single scalar value, the result will be a 1x1 array.
            object[][] result = sheet.CalculateArrayFormula(arrayFormula, calcOptions);

            // Output the aggregated sum
            Console.WriteLine("Aggregated sum of A1:B4 = " + result[0][0]);

            // -----------------------------------------------------------------
            // Demonstrate the overload that limits the size of the returned array.
            // For a formula that returns a multi‑cell array (e.g., SEQUENCE), we can
            // specify the maximum rows and columns we are interested in.
            // -----------------------------------------------------------------
            string seqFormula = "=SEQUENCE(5,2)"; // Generates a 5‑row by 2‑column array
            object[][] seqResult = sheet.CalculateArrayFormula(seqFormula, calcOptions, 5, 2);

            Console.WriteLine("\nSEQUENCE(5,2) result:");
            for (int r = 0; r < seqResult.Length; r++)
            {
                for (int c = 0; c < seqResult[r].Length; c++)
                {
                    Console.Write(seqResult[r][c] + "\t");
                }
                Console.WriteLine();
            }

            // Save the workbook (optional, demonstrates lifecycle compliance)
            workbook.Save("ArrayFormulaDemo.xlsx");
        }
    }
}
