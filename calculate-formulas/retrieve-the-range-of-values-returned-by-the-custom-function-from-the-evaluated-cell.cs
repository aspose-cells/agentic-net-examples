// Title: How to retrieve and iterate a Range object returned by a custom Excel function using Aspose.Cells in C#
// AI Prompts: Provide C# code that runs workbook.CalculateFormula with a custom engine and then casts the user‑defined function result to Aspose.Cells.Range. | Show how to loop through each cell of an Aspose.Cells.Range returned by a custom function and output the cell address and value. | Explain how to access the Address property of a Range object supplied by a custom calculation engine in Aspose.Cells.
// Common Searches: aspnet aspose.cells retrieve range from user defined function after CalculateFormula | c# get Aspose.Cells.Range object returned by custom Excel function | how to iterate cells of a range returned by a custom calculation engine in Aspose.Cells | example of custom calculation engine returning A1:B2 range in Aspose.Cells C# | extract address and values from range returned by MYRANGE function using Aspose.Cells
// Tags: custom calculation engine returning Aspose.Cells.Range | cast custom function result to Aspose.Cells.Range C# | iterate cells in Aspose.Cells.Range | retrieve range address Aspose.Cells | user-defined Excel function range output Aspose.Cells

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace CustomFunctionRangeDemo
{
    // The example creates a workbook, fills cells A1:B2 with numbers, assigns the formula =MYRANGE() to C1, and uses a custom calculation engine (MyRangeEngine) that returns the range A1:B2 as the function result. After calling workbook.CalculateFormula, the code retrieves the result, casts it to Aspose.Cells.Range, prints the range address and each cell's name and value, and saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data that will be returned by the custom function
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["B1"].PutValue(30);
                cells["B2"].PutValue(40);

                // Set a formula that calls the custom function MYRANGE()
                Cell resultCell = cells["C1"];
                resultCell.Formula = "=MYRANGE()";

                // Configure calculation options to use a custom engine
                CalculationOptions options = new CalculationOptions
                {
                    CustomEngine = new MyRangeEngine()
                };

                // Perform calculation – the custom engine will supply a Range as the result
                workbook.CalculateFormula(options);

                // Retrieve the value returned by the custom function
                object value = resultCell.Value;

                // The custom function returns a Range object, so cast accordingly
                if (value is AsposeRange returnedRange)
                {
                    Console.WriteLine("Custom function returned a range:");
                    Console.WriteLine($"Address: {returnedRange.Address}");
                    Console.WriteLine("Values in the range:");

                    // Iterate through the cells in the returned range and display their values
                    foreach (Cell cell in returnedRange)
                    {
                        Console.WriteLine($"{cell.Name}: {cell.Value}");
                    }
                }
                else
                {
                    Console.WriteLine("Custom function did not return a range. Value: " + value);
                }

                // Save the workbook (optional, demonstrates lifecycle usage)
                workbook.Save("CustomFunctionRangeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Custom calculation engine that returns a Range object for the function MYRANGE
        private class MyRangeEngine : AbstractCalculationEngine
        {
            public override void Calculate(CalculationData data)
            {
                // Ensure we are handling the expected custom function
                if (!string.IsNullOrEmpty(data.FunctionName) &&
                    data.FunctionName.Equals("MYRANGE", StringComparison.OrdinalIgnoreCase))
                {
                    // Obtain the worksheet where the function is being evaluated
                    Worksheet ws = data.Worksheet;

                    // Create a range that we want the function to return (A1:B2 in this example)
                    AsposeRange rangeToReturn = ws.Cells.CreateRange("A1:B2");

                    // Assign the range to the CalculatedValue property
                    data.CalculatedValue = rangeToReturn;
                }
            }
        }
    }
}
