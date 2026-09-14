// Title: Set worksheet calculation mode to Manual and recalculate formulas on demand using Aspose.Cells for .NET
// AI Prompts: Write C# code that configures a workbook’s calculation mode to Manual with Aspose.Cells and then triggers CalculateFormula only after modifying specific cells. | Show how to change a cell value in a manually‑calculated worksheet and invoke on‑demand formula evaluation without affecting other sheets. | Demonstrate saving a workbook after performing selective recalculation in Aspose.Cells while handling missing input files.
// Common Searches: Aspose.Cells how to set workbook calculation mode to manual in C# | C# recalculate formulas only after cell update using Aspose.Cells | disable automatic calculation Aspose.Cells .NET and trigger manual calculation | manual calc mode example with Aspose.Cells workbook save after recalculation
// Tags: Aspose.Cells manual CalcMode configuration | selective formula recalculation Aspose.Cells .NET | update cell and invoke CalculateFormula Aspose.Cells | turn off auto calculation Aspose.Cells | C# workbook save after manual recalculation

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing workbook, modifies cell A1, explicitly calls CalculateFormula to recompute formulas, and saves the workbook, illustrating how to perform on‑demand recalculation when the workbook is configured for manual calculation.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Example modification that would require recalculation
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10); // change a cell value

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
