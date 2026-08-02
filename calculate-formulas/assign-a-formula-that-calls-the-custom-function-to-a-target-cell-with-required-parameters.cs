using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Simple custom function definition (can be extended if needed)
    public class MyCustomFunctionDefinition : CustomFunctionDefinition
    {
        // Override if specific parameters must be calculated in array mode
        // public override int[] GetArrayModeParameters(string functionName)
        // {
        //     return new int[] { 0 }; // example: first parameter in array mode
        // }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data that will be used as parameters for the custom function
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);

            // Create a custom function definition instance
            MyCustomFunctionDefinition customDef = new MyCustomFunctionDefinition();

            // Set up formula parse options to use the custom function definition
            FormulaParseOptions parseOptions = new FormulaParseOptions
            {
                CustomFunctionDefinition = customDef,
                Parse = true               // enable parsing of the formula string
            };

            // Define the custom function formula
            string customFormula = "=MY_CUSTOM_FUNCTION(A1, B1)";

            // Assign the formula to the target cell (C1) using SetFormula with parse options
            // The third parameter (value) is set to null because we want Aspose.Cells to calculate it later
            sheet.Cells["C1"].SetFormula(customFormula, parseOptions, null);

            // Calculate all formulas in the workbook (including the custom one if a custom engine is provided)
            workbook.CalculateFormula();

            // Optionally, display the calculated result in the console
            Console.WriteLine("Result of custom function in C1: " + sheet.Cells["C1"].Value);

            // Save the workbook to a file
            workbook.Save("CustomFunctionDemo.xlsx");
        }
    }
}