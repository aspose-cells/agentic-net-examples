// Title: Aspose.Cells for .NET – Create a Custom Array Function to Sum a Cell Range
// Description: Demonstrates how to define a custom function (MYFUNC) with array‑mode support, register it in Aspose.Cells, apply it via an array formula ("=MYFUNC(A1:A5)"), and calculate the sum of a range so the result spills across the target cells.
// Keywords: Aspose.Cells custom function | array mode Aspose.Cells | C# spreadsheet custom engine | .NET Aspose.Cells example | array formula Aspose.Cells | sum range custom function | custom calculation engine | Excel custom function C# | GitHub Aspose.Cells demo | spreadsheet aggregation function
// Common Searches: Aspose.Cells custom array function C# | how to sum a range with a custom function in Aspose.Cells | register custom calculation engine Aspose.Cells .NET | set array formula programmatically Aspose.Cells | custom function array mode example
// Developer Intent: Implement and use a custom array‑mode function that aggregates a cell range (e.g., sum) via an array formula in Aspose.Cells for .NET.
// Use Cases: Calculate totals, averages, or other aggregates on a column without using built‑in Excel functions. | Apply the same custom aggregation across multiple worksheets by reusing the registered function. | Extend the engine to perform complex calculations (e.g., weighted sums, custom statistical metrics) on array‑mode inputs.
// AI Prompts: Generate C# code that registers a custom Aspose.Cells function to compute the product of values in an array‑mode range. | Show how to modify MyCustomEngine to return the average instead of the sum for the supplied array. | Provide an example of applying the custom function to a dynamic named range with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomArrayFunctionDemo
{
    // Custom function definition tells the engine which parameters must be evaluated in array mode.
    // Demonstrates how to define a custom function (MYFUNC) with array‑mode support, register it in Aspose.Cells, apply it via an array formula ("=MYFUNC(A1:A5)"), and calculate the sum of a range so the result spills across the target cells.
    class MyCustomFunctionDefinition : CustomFunctionDefinition
    {
        // The first (index 0) parameter of MYFUNC should be calculated in array mode.
        public override int[] GetArrayModeParameters(string functionName)
        {
            return functionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase) ? new int[] { 0 } : null;
        }
    }

    // Custom calculation engine implements the actual logic of MYFUNC.
    class MyCustomEngine : AbstractCalculationEngine
    {
        // The engine requires array‑mode values for its parameters.
        public override bool IsParamArrayModeRequired => true;

        public override void Calculate(CalculationData data)
        {
            if (!data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
                return; // Let Aspose handle other functions.

            // Retrieve the first parameter as an array (already evaluated in array mode).
            // maxRowCount / maxColumnCount = 0 means use the actual size.
            object[][] paramArray = data.GetParamValueInArrayMode(0, 0, 0);

            double sum = 0;
            foreach (object[] row in paramArray)
            {
                foreach (object item in row)
                {
                    if (item != null && double.TryParse(item.ToString(), out double d))
                        sum += d;
                }
            }

            // Return the sum as the function result.
            data.CalculatedValue = sum;
        }

        // No special forced recalculation logic needed.
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and fill sample data in column A.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            for (int i = 0; i < 5; i++)
                ws.Cells[i, 0].PutValue(i + 1); // A1..A5 = 1,2,3,4,5

            // 2. Register the custom function definition.
            wb.UpdateCustomFunctionDefinition(new MyCustomFunctionDefinition());

            // 3. Set an array formula that uses the custom function.
            // The formula will spill over 5 rows (same size as the input range).
            string arrayFormula = "=MYFUNC(A1:A5)";
            ws.Cells["B1"].SetArrayFormula(arrayFormula, 5, 1, new FormulaParseOptions());

            // 4. Prepare calculation options with the custom engine.
            CalculationOptions opts = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine()
            };

            // 5. Calculate all formulas.
            wb.CalculateFormula(opts);

            // 6. Output the result of the array formula (same value in each spilled cell).
            Console.WriteLine("Result of MYFUNC over A1:A5:");
            for (int i = 0; i < 5; i++)
                Console.WriteLine($"B{i + 1} = {ws.Cells[i, 1].Value}");

            // 7. Save the workbook (optional).
            wb.Save("CustomArrayFunctionDemo.xlsx");
        }
    }
}
