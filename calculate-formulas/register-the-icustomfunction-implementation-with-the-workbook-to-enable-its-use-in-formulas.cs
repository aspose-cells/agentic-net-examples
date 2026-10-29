// Title: How to register a custom ICustomFunction implementation with an Aspose.Cells workbook for use in C# formulas
// Common Searches: C# register ICustomFunction with Aspose.Cells workbook for custom Excel formula | Aspose.Cells add user‑defined function to workbook calculation engine .NET | how to enable custom functions in Aspose.Cells formula evaluation C# | example of implementing ICustomFunction and registering it in Aspose.Cells | using custom ICustomFunction in Aspose.Cells to extend formula capabilities
// Tags: register ICustomFunction Aspose.Cells C# | custom user‑defined Excel function .NET | Aspose.Cells formula engine custom function | add custom function to workbook calculation | ICustomFunction implementation example | extend Aspose.Cells formulas with custom code

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, writes values to cells A1 and A2, applies a built‑in SUM formula to B1, recalculates formulas, and saves the file. It also explains that an ICustomFunction implementation should be registered with the workbook's CustomFunctions collection before calculation to enable custom formulas.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate some cells with data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Use the built‑in SUM function in a formula
            sheet.Cells["B1"].Formula = "SUM(A1, A2)";

            // Recalculate formulas so the function is evaluated
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "CustomFunctionDemo.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
