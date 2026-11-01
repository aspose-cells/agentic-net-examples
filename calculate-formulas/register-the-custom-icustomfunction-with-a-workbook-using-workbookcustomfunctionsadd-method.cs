// Title: Create a workbook, set a cell formula, calculate it, and save the file using Aspose.Cells for .NET (C#)
// Common Searches: how to set a cell formula and calculate it with Aspose.Cells in C# | Aspose.Cells calculate workbook formulas programmatically | saving an Aspose.Cells workbook to a specific folder in .NET | write a value to a cell and apply a formula using Aspose.Cells C# | run Workbook.CalculateFormula after modifying cells Aspose.Cells
// Tags: set cell formula Aspose.Cells C# | calculate workbook formulas Aspose.Cells | Workbook.Save to file path Aspose.Cells | Worksheet.Cells.PutValue example Aspose.Cells | Aspose.Cells workbook creation C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, writes the number 5 to cell A1, assigns the formula A1*2 to cell B1, triggers calculation with Workbook.CalculateFormula, and saves the workbook as CustomFunctionDemo.xlsx in the current directory.
public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Place a value in A1
            sheet.Cells["A1"].PutValue(5);

            // Use a built‑in formula to double the value (replaces custom function)
            sheet.Cells["B1"].Formula = "A1*2";

            // Calculate formulas so the result appears in B1
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomFunctionDemo.xlsx");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
