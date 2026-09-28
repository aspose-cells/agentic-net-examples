// Title: Calculate all formulas in an Aspose.Cells .NET workbook using a custom MYADD user‑defined function
// AI Prompts: Create a C# custom calculation engine that implements a MYADD function and attach it to a Workbook via CalculationOptions to evaluate all formulas. | Invoke workbook.CalculateFormula with the custom engine and read the MYADD result from the target cell. | Extend the sample to register several user‑defined functions, recalculate the workbook, and then save the file.
// Common Searches: how to register a custom calculation engine in Aspose.Cells C# | using user defined functions with Aspose.Cells CalculateFormula method | evaluate MYADD function in an Excel workbook with Aspose.Cells .NET | C# code to calculate all formulas after adding custom functions in Aspose.Cells | Aspose.Cells custom engine example for Excel formula calculation
// Tags: Aspose.Cells formula calculation via custom engine | set custom engine in Aspose.Cells options | C# MYADD user-defined Excel function example | programmatic evaluation of workbook formulas | save workbook after custom formula processing

using System;
using Aspose.Cells;

// Custom calculation engine that implements a user‑defined function MYADD
// // Demonstrates creating a Workbook, inserting values, defining a CustomEngine that implements the MYADD function, attaching it via CalculationOptions, calling workbook.CalculateFormula to evaluate the formula, outputting the result, and saving the workbook.
class CustomEngine : AbstractCalculationEngine
{
    public override void Calculate(CalculationData data)
    {
        // Handle the custom function MYADD
        if (data.FunctionName.Equals("MYADD", StringComparison.OrdinalIgnoreCase))
        {
            double sum = 0;

            // Sum all numeric parameters passed to the function
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);
                if (param is double d)
                    sum += d;
                else if (param is int iVal)
                    sum += iVal;
            }

            // Set the calculated value that will be returned to the cell
            data.CalculatedValue = sum;
        }
    }
}

class Program
{
    static void Main()
    {
        // ---------- Create ----------
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);

        // Use the custom function in a formula
        sheet.Cells["B1"].Formula = "=MYADD(A1,A2)";

        // ---------- Configure calculation ----------
        // Create calculation options and attach the custom engine
        CalculationOptions options = new CalculationOptions
        {
            CustomEngine = new CustomEngine()
        };

        // ---------- Calculate ----------
        // Evaluate all formulas in the workbook using the custom engine
        workbook.CalculateFormula(options);

        // Display the result of the custom function
        Console.WriteLine("Result of MYADD(A1,A2): " + sheet.Cells["B1"].Value);

        // ---------- Save ----------
        // Save the workbook to a file
        workbook.Save("CustomEngineResult.xlsx");
    }
}
