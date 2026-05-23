using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom calculation engine that implements a function returning the current user name
    public class UserNameEngine : AbstractCalculationEngine
    {
        // Called for each custom function encountered during calculation
        public override void Calculate(CalculationData data)
        {
            // Check if the function name matches our custom function (case‑insensitive)
            if (data.FunctionName.Equals("GETUSERNAME", StringComparison.OrdinalIgnoreCase))
            {
                // Return the Windows user name as the calculated value
                data.CalculatedValue = Environment.UserName;
            }
        }

        // Ensure the function is recalculated each time (useful for audit logs)
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("GETUSERNAME", StringComparison.OrdinalIgnoreCase);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Place the custom function in cell A1
            cells["A1"].Formula = "=GETUSERNAME()";

            // Set up calculation options to use our custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new UserNameEngine()
            };

            // Perform calculation
            workbook.CalculateFormula(options);

            // Retrieve the result (the user name)
            string userName = cells["A1"].StringValue;

            // Write the audit log entry to cell B1
            cells["B1"].PutValue($"Audit: Formula evaluated by user '{userName}' at {DateTime.Now}");

            // Optional: demonstrate detection of custom functions
            Console.WriteLine($"Cell A1 has custom function: {cells["A1"].HasCustomFunction}");
            Console.WriteLine($"Workbook has custom function: {workbook.HasCustomFunction}");
            Console.WriteLine($"Calculated user name: {userName}");

            // Save the workbook (using the standard save rule)
            workbook.Save("UserNameCustomFunction.xlsx", SaveFormat.Xlsx);
        }
    }
}