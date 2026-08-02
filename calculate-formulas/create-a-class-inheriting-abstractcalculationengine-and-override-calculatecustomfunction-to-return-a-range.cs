// Title: C# – Return a Cell Range from a Custom GETRANGE() Function by Extending AbstractCalculationEngine in Aspose.Cells
// Description: Demonstrates how to inherit AbstractCalculationEngine, override Calculate to detect a GETRANGE() call, create a Range (A1:B2) on the same worksheet, assign it to CalculatedValue, and force recalculation. The sample builds a workbook, fills sample data, sets =GETRANGE() in C1, configures CalculationOptions.CustomEngine, runs the calculation, extracts the Range result, prints its address, and saves the file.
// Keywords: Aspose.Cells custom function | AbstractCalculationEngine C# | return Range from formula | GETRANGE custom engine | Aspose.Cells calculation options | C# Excel library custom formula | global .NET spreadsheet API
// Common Searches: how to create a custom function that returns a range in Aspose.Cells | override AbstractCalculationEngine to return a cell range | use CalculationOptions.CustomEngine with Aspose.Cells C# | retrieve Range object from a custom Excel formula | Aspose.Cells GETRANGE example
// Developer Intent: Implement a custom calculation engine that evaluates GETRANGE() and returns a cell Range object.
// Use Cases: Expose a predefined block of cells as a reusable range for other formulas or charts. | Provide a dynamic data source for pivot tables, data validation, or conditional formatting via a custom function. | Create a volatile custom function that always recalculates and reflects the current worksheet range.
// AI Prompts: Generate C# code that adds a GETRANGE() custom function to an Aspose.Cells workbook using AbstractCalculationEngine and returns a Range object. | Show how to modify MyRangeEngine so the function accepts start and end cell addresses as parameters and returns the corresponding Range. | Write a unit test in C# that verifies the GETRANGE custom function returns the expected A1:B2 range.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to inherit AbstractCalculationEngine, override Calculate to detect a GETRANGE() call, create a Range (A1:B2) on the same worksheet, assign it to CalculatedValue, and force recalculation. The sample builds a workbook, fills sample data, sets =GETRANGE() in C1, configures CalculationOptions.CustomEngine, runs the calculation, extracts the Range result, prints its address, and saves the file.
public class MyRangeEngine : AbstractCalculationEngine
{
    // Override Calculate to handle custom function GETRANGE
    public override void Calculate(CalculationData data)
    {
        // Check if the function being evaluated is our custom function
        if (data.FunctionName.Equals("GETRANGE", StringComparison.OrdinalIgnoreCase))
        {
            // Create a Range object that refers to cells A1:B2 on the same worksheet
            AsposeRange returnedRange = data.Worksheet.Cells.CreateRange("A1:B2");

            // Set the CalculatedValue to the Range; this will be the result of the function
            data.CalculatedValue = returnedRange;
        }
    }

    // Ensure the function is recalculated each time (useful for volatile functions)
    public override bool ForceRecalculate(string functionName)
    {
        return functionName.Equals("GETRANGE", StringComparison.OrdinalIgnoreCase);
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data that the custom function will return as a range
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].PutValue(30);
            sheet.Cells["B2"].PutValue(40);

            // Set a formula that invokes the custom function GETRANGE()
            sheet.Cells["C1"].Formula = "=GETRANGE()";

            // Configure calculation options to use our custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyRangeEngine()
            };

            // Perform calculation; the custom engine will be called
            workbook.CalculateFormula(options);

            // Retrieve the result from the cell; it should be a Range object
            object result = sheet.Cells["C1"].Value;

            if (result is AsposeRange rangeResult)
            {
                // Output the address of the returned range
                Console.WriteLine("Custom function GETRANGE returned range: " + rangeResult.RefersTo);
            }
            else
            {
                Console.WriteLine("Unexpected result type: " + (result?.GetType().Name ?? "null"));
            }

            // Save the workbook
            string outputPath = "CustomRangeFunctionDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
