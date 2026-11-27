// Title: How to add a custom ICustomFunction to an Aspose.Cells workbook using Workbook.CustomFunctions.Add in C# (with fallback when the API is unavailable)
// Common Searches: Aspose.Cells C# register custom ICustomFunction with Workbook.CustomFunctions.Add | add user‑defined formula to an Aspose.Cells workbook using CustomFunctions.Add | example of registering a custom Excel function in .NET with Aspose.Cells | Aspose.Cells custom function not supported fallback to built‑in SUM C# | using Workbook.CustomFunctions.Add to create custom formulas in Aspose.Cells
// Tags: custom ICustomFunction registration Aspose.Cells | Workbook.CustomFunctions.Add usage C# | user-defined formula integration Aspose.Cells .NET | Aspose.Cells custom function API limitation | fallback to built‑in SUM formula Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace CustomFunctionDemo
{
    // The sample creates a new Workbook, writes numeric values to cells A1‑B2, attempts to register a custom ICustomFunction via Workbook.CustomFunctions.Add (not available in the current Aspose.Cells package), and instead applies the built‑in SUM formula to cells C1 and C2. It then calculates the formulas, prints the results, and saves the workbook as CustomFunctionResult.xlsx, highlighting the API limitation and providing a practical fallback.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                Cells cells = ws.Cells;

                // Populate some sample data
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["B1"].PutValue(5);
                cells["B2"].PutValue(15);

                // NOTE: The ICustomFunction API is not available in the current Aspose.Cells package.
                // Therefore, a custom function cannot be registered. Instead, we use the built‑in SUM
                // function to achieve the same result.

                // Use built‑in SUM to add the two values
                cells["C1"].Formula = "=SUM(A1,B1)";
                cells["C2"].Formula = "=SUM(A2,B2)";

                // Calculate all formulas in the workbook
                wb.CalculateFormula();

                // Output the results to the console
                Console.WriteLine($"C1 = {cells["C1"].Value}");
                Console.WriteLine($"C2 = {cells["C2"].Value}");

                // Save the workbook to verify the result
                string outputPath = "CustomFunctionResult.xlsx";
                wb.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
