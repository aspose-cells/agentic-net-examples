// Title: How to assign a custom Excel function (MYFUNC) to a cell using Aspose.Cells for .NET with parameters
// AI Prompts: Generate C# code that creates a workbook, defines a MyCustomFunctionDefinition, and sets the formula =MYFUNC(A1,B1) in cell C1 using FormulaParseOptions. | Show how to configure Aspose.Cells to parse a custom function call and calculate the result after assigning the formula. | Write a snippet that outputs the calculated value of a custom function placed in a worksheet cell and saves the file.
// Common Searches: Aspose.Cells set custom function formula in C# workbook example | C# assign user-defined function to Excel cell using Aspose.Cells FormulaParseOptions | How to calculate workbook after adding custom function MYFUNC with Aspose.Cells | Using CustomFunctionDefinition to enable array mode parameters in Aspose.Cells
// Tags: assign custom function to cell Aspose.Cells | FormulaParseOptions custom function parsing | CustomFunctionDefinition array‑mode configuration | calculate workbook with user‑defined function | export workbook containing custom function result

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom function definition (optional override for array‑mode parameters)
    // Creates a workbook, defines a custom function MYFUNC with array‑mode for the first argument, assigns the formula =MYFUNC(A1,B1) to cell C1 via FormulaParseOptions, calculates the workbook, prints the result, and saves the file as CustomFunctionResult.xlsx.
    public class MyCustomFunctionDefinition : CustomFunctionDefinition
    {
        // Example: indicate that the first parameter should be calculated in array mode
        public override int[] GetArrayModeParameters(string functionName)
        {
            if (functionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
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

            // 2. Put some sample data that will be used as parameters
            cells["A1"].PutValue(5);
            cells["B1"].PutValue(10);

            // 3. Prepare formula‑parse options with the custom function definition
            FormulaParseOptions parseOptions = new FormulaParseOptions
            {
                Parse = true,                                 // enable parsing
                CustomFunctionDefinition = new MyCustomFunctionDefinition()
            };

            // 4. Assign a formula that calls the custom function to the target cell (C1)
            //    The formula string must start with '='.
            //    The third argument (value) is set to null because we want Aspose.Cells
            //    to calculate the result after we invoke the calculation engine.
            cells["C1"].SetFormula("=MYFUNC(A1,B1)", parseOptions, null);

            // 5. (Optional) If you have a custom calculation engine, configure it here.
            //    For this example we use the default engine, so no custom engine is set.
            //    CalculationOptions options = new CalculationOptions();

            // 6. Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // 7. Display the calculated result
            Console.WriteLine("Result of MYFUNC(A1,B1) in C1: " + cells["C1"].Value);

            // 8. Save the workbook (lifecycle rule: use provided save method)
            workbook.Save("CustomFunctionResult.xlsx");
        }
    }
}
