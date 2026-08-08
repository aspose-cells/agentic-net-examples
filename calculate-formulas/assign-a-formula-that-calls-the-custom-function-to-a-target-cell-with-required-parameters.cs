// Title: Assign a Custom Function Formula to a Cell with Parameters in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, set input values, define a custom function with array‑mode support, configure FormulaParseOptions, assign the formula "=MY_CUSTOM_FUNC(A1,B1)" to cell C1 using SetFormula, calculate the sheet, and save the result with Aspose.Cells for C#.
// Keywords: Aspose.Cells custom function | SetFormula C# | FormulaParseOptions | array mode custom function | assign formula to cell | Aspose.Cells .NET example | custom function definition | calculate workbook Aspose
// Common Searches: How to call a custom function from a worksheet cell using Aspose.Cells | SetFormula with custom function parameters C# | Configure FormulaParseOptions for custom functions in Aspose.Cells | Define array‑mode arguments for Aspose.Cells custom functions | Aspose.Cells example custom function MY_CUSTOM_FUNC
// Developer Intent: Apply a custom function to a target cell by passing required arguments and using FormulaParseOptions to enable proper parsing and calculation.
// Use Cases: Place input values in cells (e.g., A1, B1) and reference them in a custom function formula. | Override GetArrayModeParameters to treat specific arguments as array‑mode for performance or behavior reasons. | Use SetFormula with a CustomFunctionDefinition to embed the custom function call in a worksheet cell. | Trigger workbook.CalculateFormula to evaluate the custom function and retrieve the result programmatically.
// AI Prompts: Generate C# code that defines a custom function, sets up FormulaParseOptions, assigns "=MY_CUSTOM_FUNC(A1,B1)" to a cell with SetFormula, and calculates the workbook using Aspose.Cells. | Explain step‑by‑step how to enable array‑mode for a custom function argument in Aspose.Cells and invoke the function from a worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom function definition (optional – can be used to specify array‑mode parameters)
    // Demonstrates how to create a workbook, set input values, define a custom function with array‑mode support, configure FormulaParseOptions, assign the formula "=MY_CUSTOM_FUNC(A1,B1)" to cell C1 using SetFormula, calculate the sheet, and save the result with Aspose.Cells for C#.
    public class MyCustomFunctionDefinition : CustomFunctionDefinition
    {
        // Example: indicate that the first parameter should be calculated in array mode
        public override int[] GetArrayModeParameters(string functionName)
        {
            if (functionName.Equals("MY_CUSTOM_FUNC", StringComparison.OrdinalIgnoreCase))
                return new int[] { 0 }; // first argument
            return base.GetArrayModeParameters(functionName);
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate cells that will be used as parameters for the custom function
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);

            // 3. Prepare formula‑parse options with the custom function definition
            FormulaParseOptions parseOptions = new FormulaParseOptions
            {
                Parse = true,                                 // enable parsing
                CustomFunctionDefinition = new MyCustomFunctionDefinition()
            };

            // 4. Assign the custom‑function formula to the target cell (C1)
            //    The third argument (value) is set to null because we let Aspose.Cells calculate it later
            cells["C1"].SetFormula("=MY_CUSTOM_FUNC(A1,B1)", parseOptions, null);

            // 5. Calculate all formulas so the custom function result becomes visible
            //    (In a real scenario you would implement a custom calculation engine;
            //     here we just let the default engine evaluate the formula.)
            workbook.CalculateFormula();

            // 6. Output the calculated result to the console (for verification)
            Console.WriteLine("Result of MY_CUSTOM_FUNC in C1: " + cells["C1"].Value);

            // 7. Save the workbook
            workbook.Save("CustomFunctionDemo.xlsx");
        }
    }
}
