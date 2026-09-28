// Title: Create and register a GETUSERNAME custom function in Aspose.Cells for .NET and evaluate it with CalculateFormula
// AI Prompts: Define a class that inherits AbstractCalculationEngine, implement a GETUSERNAME function that returns Environment.UserName, and override ForceRecalculate for volatility. | Set CalculationOptions.CustomEngine to the custom engine, write "=GETUSERNAME()" into a worksheet cell, and invoke Workbook.CalculateFormula to compute the value. | Read the calculated value from the cell, output it to the console, and save the workbook as an .xlsx file.
// Common Searches: how to implement a custom Excel function in Aspose.Cells C# that returns the Windows user name | register a custom calculation engine in Aspose.Cells .NET for volatile functions | use CalculateFormula with a user-defined function in Aspose.Cells example | Aspose.Cells custom function GETUSERNAME sample code | audit log custom formula execution Aspose.Cells C#
// Tags: custom calculation engine Aspose.Cells .NET | GETUSERNAME user-defined function Aspose.Cells | CalculationOptions.CustomEngine usage | Workbook.CalculateFormula with custom engine | audit logging custom function Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom calculation engine that implements the GETUSERNAME function
    // The example defines a UserNameEngine class derived from AbstractCalculationEngine that implements a GETUSERNAME function returning the current Windows user name and forces recalculation. The engine is registered via CalculationOptions.CustomEngine, the formula =GETUSERNAME() is placed in cell A1, and Workbook.CalculateFormula computes the value. The result is printed to the console and the workbook is saved as CustomUserNameFunction.xlsx.
    public class UserNameEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Check if the function being calculated is our custom function
            if (data.FunctionName.Equals("GETUSERNAME", StringComparison.OrdinalIgnoreCase))
            {
                // Return the current Windows user name as the function result
                data.CalculatedValue = Environment.UserName;
            }
        }

        // Ensure the function is recalculated each time (optional, useful for volatile data)
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("GETUSERNAME", StringComparison.OrdinalIgnoreCase);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Place the custom function in a cell
            sheet.Cells["A1"].Formula = "=GETUSERNAME()";

            // Set calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new UserNameEngine()
            };

            // Calculate formulas using the custom engine (audit log can be captured here)
            workbook.CalculateFormula(options);

            // Output the result of the custom function
            Console.WriteLine("User name returned by GETUSERNAME(): " + sheet.Cells["A1"].Value);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("CustomUserNameFunction.xlsx", SaveFormat.Xlsx);
        }
    }
}
