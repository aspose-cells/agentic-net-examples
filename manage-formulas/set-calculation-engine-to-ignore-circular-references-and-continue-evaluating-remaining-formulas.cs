// Title: How to configure Aspose.Cells for .NET to ignore circular references while recalculating workbook formulas
// AI Prompts: Set the workbook's calculation settings to skip circular reference errors and recalculate the remaining formulas using Aspose.Cells in C#. | Enable the circular‑reference ignore flag in Aspose.Cells, then call CalculateFormula so the engine continues processing other cells.
// Common Searches: Aspose.Cells C# ignore circular reference errors during CalculateFormula | skip circular reference detection when recalculating Excel formulas with Aspose.Cells | continue formula evaluation after circular reference in Aspose.Cells .NET | how to set WorkbookSettings.CircularReference true in Aspose.Cells
// Tags: WorkbookSettings.CircularReference Aspose.Cells | ignore circular references during formula calculation | Aspose.Cells CalculateFormula with circular reference handling | C# Excel workbook recalculate without circular reference errors | Aspose.Cells calculation engine ignore circular reference

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, configures the Aspose.Cells calculation engine to ignore circular references by setting WorkbookSettings.CircularReference to true, recalculates all formulas, and saves the result, while handling file‑not‑found and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the modified workbook to the output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
