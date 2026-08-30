// Title: How to evaluate the IFNA function in Aspose.Cells for .NET using Workbook.CalculateFormula
// AI Prompts: Create a C# example that places a numeric value in A1 and an #N/A error in A2, assigns IFNA formulas to B1 and B2, runs Workbook.CalculateFormula, and prints the resulting values. | Generate code that demonstrates IFNA fallback handling in Aspose.Cells by setting cell formulas, invoking the calculation engine, and retrieving the evaluated results.
// Common Searches: aspocells c# IFNA formula example with CalculateFormula | how to handle #N/A errors using IFNA in Aspose.Cells .NET | C# workbook.CalculateFormula fallback value for NA error
// Tags: IFNA formula evaluation Aspose.Cells | Workbook.CalculateFormula with conditional fallback | fallback handling #N/A error C# | set cell formula IFNA Aspose.Cells | evaluate conditional formulas .NET

using System;
using Aspose.Cells;

namespace AsposeCellsIFNADemo
{
    // The program creates a new workbook, puts a numeric value in A1 and an #N/A error in A2, assigns IFNA formulas to B1 and B2 (returning the original value or a fallback text), calls Workbook.CalculateFormula to evaluate all formulas, and prints the results (42 and "fallback").
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Example 1: A1 contains a valid value, IFNA should return that value
            cells["A1"].PutValue(42);
            cells["B1"].Formula = "=IFNA(A1, \"fallback\")";

            // Example 2: A2 contains an #N/A error, IFNA should return the fallback text
            cells["A2"].Formula = "=NA()";               // Generates #N/A error
            cells["B2"].Formula = "=IFNA(A2, \"fallback\")";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the results
            Console.WriteLine("B1 (A1 is valid): " + cells["B1"].Value);      // Expected: 42
            Console.WriteLine("B2 (A2 is #N/A): " + cells["B2"].Value);      // Expected: fallback
        }
    }
}
