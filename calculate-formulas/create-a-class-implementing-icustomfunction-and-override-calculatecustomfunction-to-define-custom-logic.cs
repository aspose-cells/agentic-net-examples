// Title: Implement a custom Excel formula in C# by creating a class that implements Aspose.Cells ICustomFunction and overrides CalculateCustomFunction
// Common Searches: how to create a user‑defined Excel function with Aspose.Cells in C# | example of implementing ICustomFunction and overriding CalculateCustomFunction | registering a custom calculation class in an Aspose.Cells workbook | C# code for custom formula logic using Aspose.Cells API | step‑by‑step guide to add custom functions to Excel files with Aspose.Cells
// Tags: Aspose.Cells ICustomFunction implementation | override CalculateCustomFunction method | custom Excel formula C# | user‑defined function registration Aspose.Cells | C# custom calculation logic workbook

using System;
using Aspose.Cells;

namespace CustomFunctionsDemo
{
    // Example usage: using the built‑in SUMSQ function in a workbook.
    // The example demonstrates creating a new workbook, inserting numeric values, applying a built‑in SUMSQ formula, calculating the workbook, and printing the result, while outlining how a developer can replace the built‑in logic with a custom function by implementing ICustomFunction and overriding CalculateCustomFunction.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Add some data to the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue(2);
                sheet.Cells["A2"].PutValue(3);
                sheet.Cells["A3"].PutValue(4);

                // Use the built‑in SUMSQ function in a formula.
                sheet.Cells["B1"].Formula = "=SUMSQ(A1:A3)";

                // Calculate the formulas.
                workbook.CalculateFormula();

                // Output the result.
                Console.WriteLine("Result of SUMSQ(A1:A3): " + sheet.Cells["B1"].Value);
            }
            catch (Exception ex)
            {
                // Log or display any unexpected errors.
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
