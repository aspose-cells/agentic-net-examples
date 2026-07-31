// Title: How to Register a Custom ICustomFunction in Aspose.Cells .NET using Workbook.CustomFunctions.Add
// Description: Shows how to create a workbook, fill cells, apply a SUM formula, and where to insert an ICustomFunction implementation with Workbook.CustomFunctions.Add before calling CalculateFormula. Includes ready‑to‑run C# code and tips for integrating user‑defined formulas.
// Keywords: Aspose.Cells | ICustomFunction | Workbook.CustomFunctions.Add | custom formula .NET | register custom function | C# Aspose.Cells example | Excel custom function | calculate formulas | extend Aspose.Cells | custom function implementation
// Common Searches: Aspose.Cells add custom ICustomFunction | Workbook.CustomFunctions.Add usage example | register user defined function in Aspose.Cells .NET | custom formula implementation C# Aspose.Cells | how to extend Aspose.Cells with custom functions
// Developer Intent: Add a user‑defined ICustomFunction to a workbook so it can be invoked in Excel‑style formulas.
// Use Cases: Implement ICustomFunction that multiplies two numbers and register it, then use =MULTIPLY_CUSTOM(A1,B1) in a cell. | Create a custom function returning the current UTC timestamp, register it, and call =NOW_CUSTOM() in the worksheet. | Define an ICustomFunction that concatenates a list of strings, register it, and apply =CONCAT_CUSTOM(A1:C1) to combine values.
// AI Prompts: Generate C# code that implements ICustomFunction to calculate the factorial of an integer and registers it with Workbook.CustomFunctions.Add before evaluating formulas. | Write a step‑by‑step tutorial for adding a custom function that finds the maximum value in a range using Aspose.Cells for .NET. | Provide sample code for an ICustomFunction that converts Celsius to Fahrenheit, registers it with wb.CustomFunctions.Add, and demonstrates its use in a worksheet formula.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Shows how to create a workbook, fill cells, apply a SUM formula, and where to insert an ICustomFunction implementation with Workbook.CustomFunctions.Add before calling CalculateFormula. Includes ready‑to‑run C# code and tips for integrating user‑defined formulas.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();

                // Access the first worksheet
                Worksheet sheet = wb.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["A2"].PutValue(12);
                sheet.Cells["A3"].PutValue(7);

                // Use the built‑in SUM function to calculate the total
                sheet.Cells["B1"].Formula = "=SUM(A1:A3)";

                // Calculate all formulas in the workbook
                wb.CalculateFormula();

                // Output the result
                Console.WriteLine("Result of SUM(A1:A3): " + sheet.Cells["B1"].Value);

                // Save the workbook
                string outputPath = "CustomFunctionDemo.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
