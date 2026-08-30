// Title: Apply a custom array‑mode function to a column range with an array formula in Aspose.Cells for .NET
// AI Prompts: Write C# code that defines a user‑defined function, registers it with Aspose.Cells, and calls CalculateArrayFormula on a cell range. | Show how to implement a custom calculation engine in Aspose.Cells that returns the supplied array unchanged for a user‑defined function. | Demonstrate filling column A, applying the user‑defined array‑mode function, and writing the results to column B using Aspose.Cells.
// Common Searches: Aspose.Cells calculate array formula with user defined function in .NET | register user defined function for array mode Aspose.Cells | C# example using CalculateArrayFormula to apply a user defined function to a range | how to return an object array from a custom calculation engine in Aspose.Cells | apply user defined function to column A and write results to column B Aspose.Cells
// Tags: custom function array mode Aspose.Cells | CalculateArrayFormula with UDF | register custom function definition .NET | custom engine returning object array Aspose.Cells | apply UDF to column range | Aspose.Cells array formula example

using System;
using Aspose.Cells;

// The sample creates a workbook, populates column A with values 1‑5, registers a custom function (MYFUNC) whose first argument is evaluated in array mode, implements a custom calculation engine that returns the input array unchanged, calculates the array formula =MYFUNC(A1:A5) via CalculateArrayFormula, writes the resulting values to column B, and saves the workbook as CustomArrayFunctionDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate column A with sample data (1 to 5)
        for (int i = 0; i < 5; i++)
        {
            ws.Cells[i, 0].PutValue(i + 1); // Cells A1:A5
        }

        // Register a custom function definition that marks the first parameter
        // of MYFUNC to be evaluated in array mode
        wb.UpdateCustomFunctionDefinition(new MyCustomFunctionDefinition());

        // Set calculation options to use the custom engine
        CalculationOptions opts = new CalculationOptions
        {
            CustomEngine = new MyCustomEngine()
        };

        // Define an array formula that calls the custom function on the range A1:A5
        string arrayFormula = "=MYFUNC(A1:A5)";

        // Calculate the array formula; result is a two‑dimensional object array
        object[][] result = ws.CalculateArrayFormula(arrayFormula, opts);

        // Write the resulting values into column B (B1:B5)
        for (int i = 0; i < result.Length; i++)
        {
            ws.Cells[i, 1].PutValue(result[i][0]);
        }

        // Recalculate the workbook (not strictly necessary here)
        wb.CalculateFormula();

        // Output the results to the console
        Console.WriteLine("Results of custom array function (written to column B):");
        for (int i = 0; i < result.Length; i++)
        {
            Console.WriteLine($"B{i + 1} = {ws.Cells[i, 1].Value}");
        }

        // Save the workbook
        wb.Save("CustomArrayFunctionDemo.xlsx");
    }

    // Custom function definition: indicates that parameter 0 of MYFUNC
    // must be processed in array mode.
    class MyCustomFunctionDefinition : CustomFunctionDefinition
    {
        public override int[] GetArrayModeParameters(string functionName)
        {
            if (functionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
                return new int[] { 0 }; // first parameter
            return null;
        }
    }

    // Custom calculation engine that implements MYFUNC.
    // It simply returns the input array unchanged.
    class MyCustomEngine : AbstractCalculationEngine
    {
        // The engine requires parameters to be supplied in array mode.
        public override bool IsParamArrayModeRequired => true;

        public override void Calculate(CalculationData data)
        {
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve the first parameter as an array.
                // Passing 0 for maxRowCount/ColumnCount uses the actual size.
                object[][] paramArray = data.GetParamValueInArrayMode(0, 0, 0);

                // For demonstration, return the same array as the function result.
                data.CalculatedValue = paramArray;
            }
        }
    }
}
