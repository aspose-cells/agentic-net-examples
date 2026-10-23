// Title: Log each evaluated formula cell during workbook calculation with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells log formula evaluation for each cell after CalculateFormula | How to output evaluated formula values while calculating a workbook in Aspose.Cells | Record cell calculation results for performance analysis using Aspose.Cells .NET | Iterate over worksheet cells to capture formula results in Aspose.Cells C# example | Save workbook after logging formula evaluations with Aspose.Cells
// Tags: log formula evaluation Aspose.Cells .NET | iterate cells after CalculateFormula | output evaluated cell values C# | performance monitoring Aspose.Cells calculations | save workbook with logged formula results | exception handling Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, adds numeric data and a formula, runs CalculateFormula, then iterates through all cells to write each formula cell and its evaluated value to the console. It ensures the output directory exists, saves the workbook as 'output.xlsx', and includes basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example data and formula to demonstrate logging
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=A1+A2";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Log each cell that contains a formula and its evaluated value
            foreach (Cell cell in sheet.Cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    Console.WriteLine($"Cell {cell.Name} evaluated to {cell.Value}");
                }
            }

            // Ensure the output directory exists (if a directory part is present)
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
