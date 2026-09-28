// Title: Set worksheet calculation mode to Automatic and force a full formula recalculation with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook, sets each worksheet's CalculationMode to Automatic, and then calls CalculateFormula to recompute all formulas before saving. | Show how to enable automatic calculation for a worksheet in Aspose.Cells and trigger a complete formula refresh after modifying cell values. | Provide a step‑by‑step example of switching a workbook to Automatic calculation mode and forcing a full formula recalculation using Aspose.Cells in C#.
// Common Searches: Aspose.Cells how to enable automatic calculation for a worksheet in C# | force full formula recalculation after updating cells using Aspose.Cells .NET | set calculation mode to Automatic and recalculate all formulas Aspose.Cells example | C# Aspose.Cells recalculate workbook after data changes
// Tags: worksheet calculation mode automatic Aspose.Cells | force full formula recalculation Aspose.Cells | calculate formulas after data change C# | Aspose.Cells workbook calculation settings | automatic calculation with CalculateFormula .NET

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading or creating a workbook, switching its worksheets to Automatic calculation mode, invoking CalculateFormula to perform a complete formula recomputation, and then saving the workbook, with proper exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // Perform immediate calculation of all formulas
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "Result.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
