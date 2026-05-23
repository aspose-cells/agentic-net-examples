using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data that will be used as parameters for the custom function
            cells["A1"].PutValue(10);   // First parameter
            cells["B1"].PutValue(20);   // Second parameter

            // Define a custom function definition (can be extended if needed)
            CustomFunctionDefinition customFuncDef = new CustomFunctionDefinition();

            // Set up formula parse options to use the custom function definition
            FormulaParseOptions parseOptions = new FormulaParseOptions
            {
                CustomFunctionDefinition = customFuncDef,
                Parse = true               // Enable parsing of the formula string
            };

            // Target cell where the custom function formula will be placed
            Cell targetCell = cells["C1"];

            // Assign the formula that calls the custom function with required parameters
            // The formula string must start with '='
            string formula = "=MY_CUSTOM_FUNCTION(A1, B1)";
            targetCell.SetFormula(formula, parseOptions);

            // Optionally calculate the workbook to evaluate the formula (requires a custom engine if the function is not built‑in)
            // Here we just perform default calculation which will return #NAME? because the function is custom
            workbook.CalculateFormula();

            // Save the workbook (lifecycle save)
            workbook.Save("CustomFunctionDemo.xlsx");
        }
    }
}