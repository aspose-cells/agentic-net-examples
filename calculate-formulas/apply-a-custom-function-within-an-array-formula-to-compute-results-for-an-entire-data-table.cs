// Title: C# – Apply a Custom Array Function in Aspose.Cells to Sum a Range (MYFUNC)
// Description: This example shows how to create a custom function (MYFUNC) that runs in array‑mode, registers its definition, and plugs a custom calculation engine into Aspose.Cells. The engine sums all numeric values in the supplied range, and the function is invoked with an array formula like =MYFUNC(A1:A5) using Worksheet.CalculateArrayFormula. The workbook can then be saved or further processed.
// Keywords: Aspose.Cells custom function C# | custom array function Aspose.Cells | CalculateArrayFormula example | custom calculation engine Aspose.Cells | MYFUNC sum range | register custom function definition | .NET Excel custom function | array‑mode parameters Aspose.Cells | C# Aspose.Cells workbook calculation | Excel custom function array formula
// Common Searches: how to create a custom array function in Aspose.Cells .NET | register custom function definition for array formulas Aspose.Cells | Aspose.Cells custom calculation engine example C# | sum a column with a custom function using CalculateArrayFormula | Aspose.Cells array‑mode parameter handling
// Developer Intent: Implement and use a custom array‑mode function in Aspose.Cells to calculate a sum over a cell range.
// Use Cases: Define a custom function (MYFUNC) that processes its argument in array mode and returns the total of numeric cells. | Register the function definition with Workbook.UpdateCustomFunctionDefinition and attach a custom engine via CalculationOptions. | Execute the function in an array formula (e.g., =MYFUNC(A1:A5)) and retrieve the result with Worksheet.CalculateArrayFormula. | Integrate the custom calculation into existing spreadsheet workflows, including saving the workbook after evaluation.
// AI Prompts: Generate C# code for an Aspose.Cells custom function called PRODUCT that multiplies all numbers in a range using array‑mode. | Show how to register a custom function definition and a custom calculation engine, then evaluate an array formula across multiple columns in Aspose.Cells. | Explain error handling strategies inside a custom Aspose.Cells calculation engine and how to return proper Excel error values.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomArrayFunctionDemo
{
    // Custom function definition that tells the engine which parameters need array‑mode calculation
    // This example shows how to create a custom function (MYFUNC) that runs in array‑mode, registers its definition, and plugs a custom calculation engine into Aspose.Cells. The engine sums all numeric values in the supplied range, and the function is invoked with an array formula like =MYFUNC(A1:A5) using Worksheet.CalculateArrayFormula. The workbook can then be saved or further processed.
    class MyCustomFunctionDefinition : CustomFunctionDefinition
    {
        // The first (and only) parameter of MYFUNC should be calculated in array mode
        public override int[] GetArrayModeParameters(string functionName)
        {
            if (functionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
                return new int[] { 0 };   // zero‑based index of the parameter
            return null;
        }
    }

    // Custom calculation engine that implements the logic of MYFUNC
    class MyCustomEngine : AbstractCalculationEngine
    {
        // Indicate that this engine requires array‑mode parameters (optional, but safe)
        public override bool IsParamArrayModeRequired => true;

        public override void Calculate(CalculationData data)
        {
            try
            {
                if (!data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
                    return; // ignore other functions

                // Use large values to let the engine determine the actual needed size
                int maxRows = int.MaxValue;
                int maxCols = int.MaxValue;

                // Retrieve the parameter in array mode
                object[][] paramArray = data.GetParamValueInArrayMode(0, maxRows, maxCols);

                double sum = 0;
                foreach (object[] row in paramArray)
                {
                    foreach (object item in row)
                    {
                        if (item != null && double.TryParse(item.ToString(), out double d))
                            sum += d;
                    }
                }

                // Return the sum as the function result
                data.CalculatedValue = sum;
            }
            catch (Exception ex)
            {
                // In case of unexpected errors, set the result to an error value
                data.CalculatedValue = $"#ERROR: {ex.Message}";
            }
        }

        // No special forced recalculation logic
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a workbook and fill a simple data table (A1:A5)
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                for (int i = 0; i < 5; i++)
                    ws.Cells[i, 0].PutValue(i + 1);   // A1=1, A2=2, ..., A5=5

                // 2. Register the custom function definition
                wb.UpdateCustomFunctionDefinition(new MyCustomFunctionDefinition());

                // 3. Prepare calculation options with the custom engine
                CalculationOptions opts = new CalculationOptions
                {
                    CustomEngine = new MyCustomEngine()
                };

                // 4. Use an array formula that calls the custom function over the whole column
                string arrayFormula = "=MYFUNC(A1:A5)";

                // Calculate the array formula; result is a 2‑dimensional object array
                object[][] result = ws.CalculateArrayFormula(arrayFormula, opts);

                // 5. Output the computed value (should be 1+2+3+4+5 = 15)
                Console.WriteLine("Result of custom array function MYFUNC: " + result[0][0]);

                // 6. Save the workbook (optional, demonstrates lifecycle compliance)
                wb.Save("CustomArrayFunctionDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
