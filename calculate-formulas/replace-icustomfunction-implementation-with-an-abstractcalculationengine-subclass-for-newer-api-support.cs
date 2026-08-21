// Title: Aspose.Cells .NET – Replace ICustomFunction with an AbstractCalculationEngine Subclass
// Description: Demonstrates how to create a custom calculation engine by inheriting from AbstractCalculationEngine. The engine disables array‑mode and literal‑text handling, skips built‑in functions, and forces recalculation for a volatile function. It implements two sample functions – MYFUNC (adds two numbers) and MYVOLATILEFUNC (returns the current ISO‑8601 timestamp). The example shows workbook creation, formula insertion, optional CustomFunctionDefinition registration, CalculationOptions configuration with the custom engine, formula evaluation, console output, and saving the workbook as XLSX and PDF.
// Keywords: Aspose.Cells | AbstractCalculationEngine | custom calculation engine | C# | .NET | replace ICustomFunction | custom function volatile | CalculationOptions.CustomEngine | register custom function definition | Excel to PDF conversion
// Common Searches: How to use AbstractCalculationEngine in Aspose.Cells .NET | Replace ICustomFunction with AbstractCalculationEngine example | Create volatile custom function in Aspose.Cells | Set up CalculationOptions.CustomEngine for custom formulas | Define array‑mode parameters for custom functions Aspose.Cells
// Developer Intent: Migrate legacy ICustomFunction code to the newer AbstractCalculationEngine API and evaluate custom formulas within an Aspose.Cells workbook.
// Use Cases: Add a simple custom function (MYFUNC) that sums two cell values using a custom engine. | Implement a volatile custom function (MYVOLATILEFUNC) that returns the current timestamp and triggers automatic recalculation. | Specify array‑mode parameters for custom functions via a CustomFunctionDefinition (e.g., ARRAYSUM). | Calculate all workbook formulas with a custom engine and export results to XLSX and PDF.
// AI Prompts: Generate C# code that defines an AbstractCalculationEngine subclass handling a custom function that multiplies three parameters and registers it with a workbook. | Show how to configure CalculationOptions.CustomEngine to evaluate a custom function that returns a formatted date string in Aspose.Cells. | Explain the steps to make a custom function volatile so Aspose.Cells forces recalculation on each workbook open.

using System;
using Aspose.Cells;

namespace CustomEngineDemo
{
    // Custom calculation engine derived from AbstractCalculationEngine
    // Demonstrates how to create a custom calculation engine by inheriting from AbstractCalculationEngine. The engine disables array‑mode and literal‑text handling, skips built‑in functions, and forces recalculation for a volatile function. It implements two sample functions – MYFUNC (adds two numbers) and MYVOLATILEFUNC (returns the current ISO‑8601 timestamp). The example shows workbook creation, formula insertion, optional CustomFunctionDefinition registration, CalculationOptions configuration with the custom engine, formula evaluation, console output, and saving the workbook as XLSX and PDF.
    public class MyCustomEngine : AbstractCalculationEngine
    {
        // This engine does not need array mode or literal text
        public override bool IsParamArrayModeRequired => false;
        public override bool IsParamLiteralRequired => false;

        // Do not process built‑in functions
        public override bool ProcessBuiltInFunctions => false;

        // Force recalculation for a volatile custom function (optional)
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("MYVOLATILEFUNC", StringComparison.OrdinalIgnoreCase);
        }

        // Core calculation logic for custom functions
        public override void Calculate(CalculationData data)
        {
            // Helper to obtain a numeric value from a parameter (handles ranges)
            double GetNumericParam(int index)
            {
                object param = data.GetParamValue(index);
                if (param == null)
                    return 0.0;

                // If the parameter is a range represented as a 2‑D array, take the first element
                if (param is double[,] dArr && dArr.Length > 0)
                    return dArr[0, 0];

                if (param is object[,] oArr && oArr.Length > 0)
                    return Convert.ToDouble(oArr[0, 0] ?? 0);

                // Fallback: try direct conversion
                return Convert.ToDouble(param);
            }

            // Handle a custom function named MYFUNC that adds two numbers
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                double val0 = GetNumericParam(0);
                double val1 = GetNumericParam(1);
                data.CalculatedValue = val0 + val1;
                return;
            }

            // Handle a custom volatile function that returns current timestamp
            if (data.FunctionName.Equals("MYVOLATILEFUNC", StringComparison.OrdinalIgnoreCase))
            {
                data.CalculatedValue = DateTime.Now.ToString("O");
                return;
            }

            // For any other function, let the default engine handle it (do nothing)
        }
    }

    // Optional: custom function definition to mark parameters that need array mode
    public class MyCustomFunctionDefinition : CustomFunctionDefinition
    {
        // Example: for a function named ARRAYSUM, the first parameter should be evaluated in array mode
        public override int[] GetArrayModeParameters(string functionName)
        {
            if (functionName.Equals("ARRAYSUM", StringComparison.OrdinalIgnoreCase))
                return new int[] { 0 }; // first parameter
            return base.GetArrayModeParameters(functionName);
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with sample data
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // Use the custom function MYFUNC (adds two numbers)
                sheet.Cells["B1"].Formula = "=MYFUNC(A1, A2)";

                // Use a volatile custom function that returns the current timestamp
                sheet.Cells["B2"].Formula = "=MYVOLATILEFUNC()";

                // Set up calculation options with the custom engine
                CalculationOptions options = new CalculationOptions
                {
                    CustomEngine = new MyCustomEngine(),
                    IgnoreError = false,
                    Recursive = true
                };

                // Register custom function definition (if needed)
                workbook.UpdateCustomFunctionDefinition(new MyCustomFunctionDefinition());

                // Calculate all formulas using the custom engine
                workbook.CalculateFormula(options);

                // Output results to console
                Console.WriteLine("Result of MYFUNC(A1, A2): " + sheet.Cells["B1"].Value);
                Console.WriteLine("Result of MYVOLATILEFUNC(): " + sheet.Cells["B2"].Value);

                // Save the workbook (both Excel and PDF formats)
                workbook.Save("CustomEngineDemo.xlsx");
                workbook.Save("CustomEngineDemo.pdf");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
