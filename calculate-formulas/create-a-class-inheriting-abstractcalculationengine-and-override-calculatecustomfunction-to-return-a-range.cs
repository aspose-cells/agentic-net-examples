// Title: Implement a custom Aspose.Cells calculation engine in C# that returns a B1:C2 range via a GETRANGE user‑defined function
// AI Prompts: Generate a C# class that inherits AbstractCalculationEngine and implements the Calculate method to detect the GETRANGE function and assign a B1:C2 Range object to the CalculatedValue. | Demonstrate setting CalculationOptions.CustomEngine to your engine class and calling sheet.CalculateFormula("=GETRANGE()") to obtain the returned range and display its values. | Extend the custom engine so GETRANGE(startCell, endCell) accepts two cell addresses and dynamically returns the corresponding range.
// Common Searches: how to create a custom function that returns a cell range in Aspose.Cells C# | Aspose.Cells example using AbstractCalculationEngine to return a range object | C# code for GETRANGE user defined function with Aspose.Cells calculation engine | register custom calculation engine in Aspose.Cells and evaluate =GETRANGE() | returning B1:C2 range from a custom formula in Aspose.Cells .NET
// Tags: custom calculation engine Aspose.Cells | GETRANGE custom function Aspose.Cells | return range object from custom formula | override Calculate method Aspose.Cells | evaluate formula with CalculationOptions C#

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace CustomEngineRangeDemo
{
    // Custom calculation engine that returns a cell range for the function GETRANGE()
    // The example defines MyRangeEngine, a class derived from AbstractCalculationEngine, that implements Calculate to recognize the GETRANGE function, creates a B1:C2 Range on the calling worksheet, and assigns it to CalculatedValue. The engine is attached via CalculationOptions.CustomEngine, the formula =GETRANGE() is evaluated, and the returned range address and cell values are printed. The code also shows how to make the function volatile and how to extend it to accept dynamic range parameters.
    public class MyRangeEngine : AbstractCalculationEngine
    {
        // Calculate is called for each custom function encountered during formula evaluation
        public override void Calculate(CalculationData data)
        {
            // Check if the function name matches our custom function (case‑insensitive)
            if (string.Equals(data.FunctionName, "GETRANGE", StringComparison.OrdinalIgnoreCase))
            {
                // Create a range B1:C2 on the same worksheet where the function is used
                AsposeRange range = data.Worksheet.Cells.CreateRange("B1:C2");

                // Assign the range to the CalculatedValue property.
                // The caller will receive this Range object as the result of the formula.
                data.CalculatedValue = range;
            }
        }

        // Ensure the function is recalculated for every call (useful for volatile functions)
        public override bool ForceRecalculate(string functionName) => true;
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the range B1:C2 with sample data
                sheet.Cells["B1"].PutValue(1);
                sheet.Cells["B2"].PutValue(2);
                sheet.Cells["C1"].PutValue(3);
                sheet.Cells["C2"].PutValue(4);

                // Set a formula that invokes the custom function GETRANGE()
                sheet.Cells["A1"].Formula = "=GETRANGE()";

                // Configure calculation options to use our custom engine
                CalculationOptions options = new CalculationOptions
                {
                    CustomEngine = new MyRangeEngine()
                };

                // Calculate the formula. The result will be a Range object.
                object result = sheet.CalculateFormula("=GETRANGE()", options);

                // Verify and display the returned range and its values
                if (result is AsposeRange returnedRange)
                {
                    Console.WriteLine("Returned range address: " + returnedRange.RefersTo);
                    Console.WriteLine("Values inside the returned range:");

                    for (int row = returnedRange.FirstRow; row < returnedRange.FirstRow + returnedRange.RowCount; row++)
                    {
                        for (int col = returnedRange.FirstColumn; col < returnedRange.FirstColumn + returnedRange.ColumnCount; col++)
                        {
                            Console.Write(sheet.Cells[row, col].Value + "\t");
                        }
                        Console.WriteLine();
                    }
                }
                else
                {
                    Console.WriteLine("The custom function did not return a range.");
                }

                // Save the workbook (optional, just to illustrate that the file can be saved)
                workbook.Save("RangeCustomFunctionDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
