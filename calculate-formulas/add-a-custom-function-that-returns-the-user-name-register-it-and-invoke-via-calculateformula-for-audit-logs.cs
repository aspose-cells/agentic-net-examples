// Title: C# Custom USERNAME() Function in Aspose.Cells – Register Engine, CalculateFormula & Audit Log
// Description: Shows how to extend AbstractCalculationEngine with a USERNAME() function, register it via CalculationOptions, evaluate it using Workbook.CalculateFormula, and write an audit entry for each call.
// Keywords: Aspose.Cells | C# custom function | USERNAME() | AbstractCalculationEngine | CalculationOptions | CalculateFormula | audit log | volatile function | Environment.UserName | Excel automation | spreadsheet custom engine
// Common Searches: Aspose.Cells custom USERNAME function C# | register custom calculation engine Aspose.Cells | calculate formula with custom engine .NET | log custom function evaluation Aspose.Cells | volatile custom functions Aspose.Cells | get Windows user name in spreadsheet using Aspose.Cells
// Developer Intent: Implement a USERNAME() function that returns the current Windows user, register it as a custom calculation engine, run workbook.CalculateFormula, and capture each evaluation in an audit log.
// Use Cases: Insert the logged‑in Windows user into generated reports. | Maintain an audit trail of spreadsheet calculations for compliance. | Force recalculation of volatile data such as usernames or timestamps. | Extend the engine to support additional volatile functions like NOW() with logging.
// AI Prompts: Write C# code that defines a class inheriting AbstractCalculationEngine to implement a USERNAME() function that returns Environment.UserName and logs each evaluation. | Show how to attach the custom engine to a Workbook using CalculationOptions and invoke workbook.CalculateFormula to evaluate the function. | Create an example that adds another volatile function (e.g., NOW()) to the same engine with audit logging.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom calculation engine that implements the USERNAME() function
    // Shows how to extend AbstractCalculationEngine with a USERNAME() function, register it via CalculationOptions, evaluate it using Workbook.CalculateFormula, and write an audit entry for each call.
    public class UserNameEngine : AbstractCalculationEngine
    {
        // Called for each custom function during calculation
        public override void Calculate(CalculationData data)
        {
            // Check if the function being evaluated is USERNAME (case‑insensitive)
            if (data.FunctionName.Equals("USERNAME", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve the current Windows user name
                string userName = Environment.UserName;

                // Set the calculated value – this will be returned to the cell
                data.CalculatedValue = userName;

                // Optional audit log (could be written to a file, DB, etc.)
                Console.WriteLine($"Audit Log: USERNAME() evaluated, result = '{userName}'");
            }
        }

        // Ensure the function is recalculated each time (useful for volatile data like user name)
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("USERNAME", StringComparison.OrdinalIgnoreCase);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Place the custom function in a cell
            sheet.Cells["A1"].Formula = "=USERNAME()";

            // Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new UserNameEngine()
            };

            // Perform calculation using the custom engine (invoke via CalculateFormula)
            workbook.CalculateFormula(options);

            // Output the result of the custom function
            Console.WriteLine($"Cell A1 value (user name): {sheet.Cells["A1"].Value}");

            // Save the workbook (lifecycle rule: save)
            workbook.Save("UserNameCustomFunction.xlsx", SaveFormat.Xlsx);
        }
    }
}
