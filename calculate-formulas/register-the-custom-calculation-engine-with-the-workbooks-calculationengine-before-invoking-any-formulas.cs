// Title: C# – Register a Custom Calculation Engine in Aspose.Cells Before Formula Evaluation
// Description: Shows how to build a workbook, insert values, set a formula that calls a custom function (MYADD), implement a custom engine by extending AbstractCalculationEngine, register it through CalculationOptions.CustomEngine, run wb.CalculateFormula, and save the workbook.
// Keywords: Aspose.Cells | C# | custom calculation engine | CalculationOptions.CustomEngine | AbstractCalculationEngine | custom function | MYADD | CalculateFormula | Excel automation | spreadsheet custom functions
// Common Searches: register custom calculation engine Aspose.Cells .NET | use custom functions with Aspose.Cells CalculateFormula | extend AbstractCalculationEngine example | set CalculationOptions.CustomEngine before wb.CalculateFormula | Aspose.Cells custom formula engine tutorial
// Developer Intent: Learn how to plug a user‑defined calculation engine into an Aspose.Cells workbook so that formulas containing custom functions are evaluated correctly.
// Use Cases: Add domain‑specific functions (e.g., MYADD) to Excel calculations. | Replace the default engine to handle proprietary business logic. | Calculate and persist results of workbooks that rely on custom formulas.
// AI Prompts: Generate C# code that registers a custom calculation engine in Aspose.Cells and evaluates a workbook. | Explain step‑by‑step how to implement and debug a custom function using AbstractCalculationEngine. | Show how to use CalculationOptions.CustomEngine to run formulas with user‑defined functions.

using System;
using Aspose.Cells;

namespace CustomEngineDemo
{
    // Shows how to build a workbook, insert values, set a formula that calls a custom function (MYADD), implement a custom engine by extending AbstractCalculationEngine, register it through CalculationOptions.CustomEngine, run wb.CalculateFormula, and save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate cells with sample data
            ws.Cells["A1"].PutValue(5);
            ws.Cells["A2"].PutValue(7);

            // Set a formula that uses a custom function "MYADD"
            ws.Cells["A3"].Formula = "=MYADD(A1,A2)";

            // Create calculation options and register the custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine()
            };

            // Calculate formulas using the custom engine
            wb.CalculateFormula(options);

            // Display the result of the custom function
            Console.WriteLine("Result of MYADD: " + ws.Cells["A3"].Value);

            // Save the workbook
            wb.Save("CustomEngineResult.xlsx");
        }
    }

    // Custom calculation engine extending the default engine
    public class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle the custom function "MYADD"
            if (data.FunctionName.Equals("MYADD", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve parameter values
                double v1 = Convert.ToDouble(data.GetParamValue(0));
                double v2 = Convert.ToDouble(data.GetParamValue(1));

                // Set the calculated result (simple addition)
                data.CalculatedValue = v1 + v2;
            }
        }
    }
}
