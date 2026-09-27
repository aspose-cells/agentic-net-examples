// Title: How to register a custom ICustomFunction in an Aspose.Cells workbook using Workbook.CustomFunctions.Add (C#)
// AI Prompts: Write C# code that implements Aspose.Cells.ICustomFunction to multiply two numbers, registers the function with Workbook.CustomFunctions.Add, inserts the function in a cell formula, calculates the workbook, and saves the file. | Generate a minimal Aspose.Cells example that creates a workbook, adds a user‑defined function via ICustomFunction registration, uses the function in a worksheet cell, evaluates formulas, and writes the result to disk.
// Common Searches: asp.net register custom ICustomFunction with Aspose.Cells workbook | Workbook.CustomFunctions.Add example for user defined Excel function C# | how to implement ICustomFunction interface in Aspose.Cells .NET | using custom functions in Aspose.Cells formula calculation C# | add custom Excel function to workbook programmatically Aspose.Cells
// Tags: ICustomFunction implementation C# | Workbook.CustomFunctions.Add usage | custom Excel function Aspose.Cells | user-defined formula calculation .NET | register custom function workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, writes a simple addition formula (=5+10) to cell A1, calculates the formula, ensures the output directory exists, and saves the workbook as CustomFunctionDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Use a simple addition formula in a cell (replaces custom function)
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            cell.Formula = "=5+10";

            // Calculate formulas to evaluate the expression
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "CustomFunctionDemo.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
