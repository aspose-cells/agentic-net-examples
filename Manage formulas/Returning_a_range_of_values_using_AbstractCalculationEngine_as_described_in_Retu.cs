using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineDemo
{
    // Custom engine that returns a 2x2 range (array) for the function RANGEFUNC()
    class RangeReturnEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Check that the invoked function is our custom one
            if (data.FunctionName != null &&
                data.FunctionName.Equals("RANGEFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Build a 2‑row by 2‑column array
                object[][] result = new object[2][];
                result[0] = new object[] { 10, 20 };
                result[1] = new object[] { 30, 40 };

                // Assign the array as the calculated value
                data.CalculatedValue = result;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Load an existing XLSX workbook ----------
            // (Replace "input.xlsx" with the actual path of your source file)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Place a formula that calls the custom function
            sheet.Cells["A1"].Formula = "=RANGEFUNC()";

            // ---------- Set up calculation options with the custom engine ----------
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new RangeReturnEngine()
            };

            // Perform calculation using the custom engine
            workbook.CalculateFormula(options);

            // ---------- Retrieve and display the returned range ----------
            // The cell A1 now holds a ReferredArea that represents the spilled array.
            // Aspose.Cells stores the array as a 2‑dimensional object array.
            object value = sheet.Cells["A1"].Value;

            if (value is object[][] arrayResult)
            {
                Console.WriteLine("RANGEFUNC returned the following array:");
                for (int r = 0; r < arrayResult.Length; r++)
                {
                    for (int c = 0; c < arrayResult[r].Length; c++)
                    {
                        Console.Write(arrayResult[r][c] + "\t");
                    }
                    Console.WriteLine();
                }
            }
            else
            {
                Console.WriteLine("Unexpected result type: " + (value?.GetType().FullName ?? "null"));
            }

            // ---------- Save the workbook ----------
            // (Replace "output.xlsx" with the desired output path)
            workbook.Save("output.xlsx");
        }
    }
}