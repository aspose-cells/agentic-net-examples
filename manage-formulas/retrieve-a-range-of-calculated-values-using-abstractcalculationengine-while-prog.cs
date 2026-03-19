using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineDemo
{
    // Custom engine that extracts a range of values passed to the custom function GETRANGE
    class RangeEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Ensure we are handling the expected custom function
            if (data.FunctionName != null && data.FunctionName.Equals("GETRANGE", StringComparison.OrdinalIgnoreCase))
            {
                // The first parameter should be a range (ReferredArea)
                object param = data.GetParamValue(0);
                if (param is ReferredArea area)
                {
                    // Retrieve all values in the area, calculating any formulas inside the area
                    object values = area.GetValues(true); // true => calculate formulas recursively

                    // Set the calculated value to the retrieved values (can be a 2‑D array)
                    data.CalculatedValue = values;
                }
                else
                {
                    // If the parameter is not a range, return an error value
                    data.CalculatedValue = "#VALUE!";
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that will be referenced by the custom function
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].PutValue(30);
            sheet.Cells["B2"].PutValue(40);

            // Place a formula that calls the custom function GETRANGE on the range A1:B2
            sheet.Cells["C1"].Formula = "=GETRANGE(A1:B2)";

            // Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new RangeEngine()
            };

            // Perform calculation using the custom engine
            workbook.CalculateFormula(options);

            // Retrieve the result from cell C1
            object result = sheet.Cells["C1"].Value;

            // The result is expected to be a 2‑D array; display its contents
            if (result is object[,] array)
            {
                Console.WriteLine("Values returned by GETRANGE:");
                for (int i = 0; i < array.GetLength(0); i++)
                {
                    for (int j = 0; j < array.GetLength(1); j++)
                    {
                        Console.Write(array[i, j] + "\t");
                    }
                    Console.WriteLine();
                }
            }
            else
            {
                Console.WriteLine("Result: " + result);
            }

            // Save the workbook (optional, to verify that calculation succeeded)
            workbook.Save("output.xlsx");
        }
    }
}