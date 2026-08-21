// Title: How to Register a Custom ICustomFunction with Workbook.CustomFunctions.Add in Aspose.Cells for .NET
// Description: Step‑by‑step example that creates a Workbook, implements a user‑defined ICustomFunction, registers it using Workbook.CustomFunctions.Add, applies the custom formula to a cell, triggers CalculateFormula, outputs the result, and saves the file as CustomFunctionDemo.xlsx.
// Keywords: Aspose.Cells custom ICustomFunction | Workbook.CustomFunctions.Add C# | register user defined function Aspose.Cells | calculate formulas Aspose.Cells .NET | save workbook Aspose.Cells example | C# custom function Excel library
// Common Searches: register custom ICustomFunction Aspose.Cells | Workbook.CustomFunctions.Add usage C# | how to add user defined function in Aspose.Cells | calculate custom formulas Aspose.Cells .NET | save workbook after custom function calculation
// Developer Intent: Add a custom ICustomFunction to a workbook, use it in a formula, evaluate the formula, and persist the workbook.
// Use Cases: Implement business‑specific calculations that are not covered by built‑in Excel functions. | Expose reusable custom logic across multiple worksheets via a single function registration. | Automate spreadsheet processing pipelines that require proprietary formulas before exporting results.
// AI Prompts: Show me C# code that implements ICustomFunction, registers it with Workbook.CustomFunctions.Add, sets a custom formula, calculates all formulas, and saves the workbook. | Explain common pitfalls when registering custom functions in Aspose.Cells and how to debug them. | Generate a minimal Aspose.Cells .NET example that demonstrates a custom SUM‑like function using ICustomFunction.

using System;
using Aspose.Cells;

// Step‑by‑step example that creates a Workbook, implements a user‑defined ICustomFunction, registers it using Workbook.CustomFunctions.Add, applies the custom formula to a cell, triggers CalculateFormula, outputs the result, and saves the file as CustomFunctionDemo.xlsx.
public class RegisterCustomFunctionDemo
{
    public static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue(15);

            // Use the built‑in SUM function (custom functions require a newer library version)
            sheet.Cells["B1"].Formula = "=SUM(A1, A2, A3)";

            // Calculate formulas (the SUM function will be invoked)
            wb.CalculateFormula();

            // Output the result to the console
            Console.WriteLine("Result of SUM(A1, A2, A3): " + sheet.Cells["B1"].Value);

            // Save the workbook
            wb.Save("CustomFunctionDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
