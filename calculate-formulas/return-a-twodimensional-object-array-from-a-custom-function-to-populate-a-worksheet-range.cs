// Title: Fill an Excel range from a C# object[][] using Aspose.Cells SetArrayFormula
// Description: Shows how to return a two‑dimensional object array in C# and assign it to a worksheet range with Cell.SetArrayFormula, including formula setup, row/column sizing, optional calculation, and file saving.
// Keywords: Aspose.Cells | SetArrayFormula | C# object[][] | pre‑calculated values | populate Excel range | array formula overload | Excel automation .NET | Workbook.CalculateFormula | Excel data injection | cell array assignment
// Common Searches: Aspose.Cells SetArrayFormula example C# | return object[][] for Excel cells | populate Excel range with pre‑calculated values | C# array formula values parameter Aspose | how to use SetArrayFormula overload
// Developer Intent: Insert a 2‑D array returned by C# code into an Excel worksheet without re‑evaluating the formula.
// Use Cases: Load lookup data from a database or file and inject it directly into a workbook for fast reporting. | Prepare large calculation results in code and write them to Excel in one call to improve performance. | Create static chart data series by supplying pre‑computed values to a range via SetArrayFormula.
// AI Prompts: Generate C# code that reads a CSV file and returns its contents as an object[][] for SetArrayFormula. | Explain how to determine the dimensions of an object[][] at runtime and call SetArrayFormula with dynamic row and column counts. | Show how to handle mixed data types (numbers, dates, strings) in the object[][] when populating an Excel range with SetArrayFormula.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Shows how to return a two‑dimensional object array in C# and assign it to a worksheet range with Cell.SetArrayFormula, including formula setup, row/column sizing, optional calculation, and file saving.
    class Program
    {
        // Custom function that returns a two‑dimensional object array.
        // In a real scenario this could be any logic that builds the data.
        static object[][] GetSampleData()
        {
            // Create a 3x2 array of sample values.
            return new object[][]
            {
                new object[] { 10, 20 },
                new object[] { 30, 40 },
                new object[] { 50, 60 }
            };
        }

        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Put some source data that the array formula will reference.
            cells["A1"].PutValue(1);
            cells["A2"].PutValue(2);
            cells["A3"].PutValue(3);
            cells["B1"].PutValue(4);
            cells["B2"].PutValue(5);
            cells["B3"].PutValue(6);

            // 3. Define the array formula. This example multiplies the range A1:B3 by 2.
            string arrayFormula = "=A1:B3*2";

            // 4. Retrieve the two‑dimensional object array from the custom function.
            object[][] preCalculatedValues = GetSampleData();

            // 5. Apply the array formula to the top‑left cell of the target range.
            //    The overload with values supplies the pre‑calculated results,
            //    so the formula does not need to be evaluated again.
            Cell targetCell = cells["C1"];
            targetCell.SetArrayFormula(
                arrayFormula,          // formula expression
                rowNumber: 3,          // number of rows in the result range
                columnNumber: 2,       // number of columns in the result range
                options: new FormulaParseOptions(), // default parsing options
                values: preCalculatedValues); // two‑dimensional array returned by the custom function

            // 6. Calculate the workbook (optional, not required when values are supplied).
            workbook.CalculateFormula();

            // 7. Save the workbook.
            workbook.Save("CustomFunctionArrayResult.xlsx");

            // 8. Output the populated range to the console for verification.
            Console.WriteLine("Populated range C1:D3:");
            for (int r = 0; r < 3; r++)
            {
                for (int c = 0; c < 2; c++)
                {
                    Console.Write(cells[0 + r, 2 + c].Value + "\t");
                }
                Console.WriteLine();
            }
        }
    }
}
