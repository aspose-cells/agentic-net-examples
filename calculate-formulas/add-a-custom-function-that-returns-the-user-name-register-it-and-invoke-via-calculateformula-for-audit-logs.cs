using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom calculation engine that implements the GETUSERNAME() function
    public class UserNameEngine : AbstractCalculationEngine
    {
        // Called for each custom function during calculation
        public override void Calculate(CalculationData data)
        {
            // Check if the function being evaluated is our custom function
            if (data.FunctionName.Equals("GETUSERNAME", StringComparison.OrdinalIgnoreCase))
            {
                // Return the current Windows user name as the function result
                data.CalculatedValue = Environment.UserName;
            }
        }

        // Ensure the function is recalculated each time (useful for volatile data like user name)
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("GETUSERNAME", StringComparison.OrdinalIgnoreCase);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Place the custom function in cell A1
            // The formula name must match the name used in the custom engine
            sheet.Cells["A1"].Formula = "=GETUSERNAME()";

            // Configure calculation options to use the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new UserNameEngine()
            };

            // Perform calculation (invokes the custom function)
            workbook.CalculateFormula(options);

            // Retrieve and display the result
            string userName = sheet.Cells["A1"].StringValue;
            Console.WriteLine("Custom function result (User Name): " + userName);

            // Audit log example – write to console (could be written to a file or DB)
            Console.WriteLine($"[Audit] GETUSERNAME() evaluated at {DateTime.UtcNow:u} – Result: {userName}");

            // Save the workbook (lifecycle: save)
            workbook.Save("UserNameFunctionDemo.xlsx");
        }
    }
}