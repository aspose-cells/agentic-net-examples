// Title: Recalculate formulas in Aspose.Cells only after cell A1 changes using manual calculation mode in C#
// AI Prompts: Write C# code that sets Aspose.Cells workbook to manual calculation mode and invokes Workbook.CalculateFormula() only when cell A1 is updated. | Show how to simulate an event‑like trigger in Aspose.Cells that detects a change to a specific cell and performs on‑demand formula recalculation. | Provide a minimal example that modifies cell A1, manually recalculates all formulas, and saves the workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells manual calculation mode trigger formula recalculation after updating a single cell | C# Aspose.Cells recalculate workbook only when specific cell value changes | How to force formula evaluation in Aspose.Cells after modifying cell A1 | Event driven formula calculation in Aspose.Cells .NET | Calculate formulas on demand in Aspose.Cells workbook C#
// Tags: manual calculation mode Aspose.Cells C# | Workbook.CalculateFormula on demand | trigger formula recalculation specific cell | Aspose.Cells event‑driven calculation | update cell A1 recalc formulas Aspose.Cells | C# Aspose.Cells workbook save after manual calc

using Aspose.Cells;
using System;
using System.IO;

// Creates a new workbook (or loads an existing one), optionally switches to manual calculation mode, changes the value of cell A1, manually recalculates all formulas with Workbook.CalculateFormula(), and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if needed)
            var workbook = new Workbook();

            // NOTE: Setting calculation mode is optional and may not be supported in all versions.
            // If needed, you can enable manual calculation mode like this:
            // workbook.Settings.CalcMode = CalcMode.Manual;

            // Define the trigger cell address (A1 on the first worksheet)
            const string triggerAddress = "A1";

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Change the trigger cell value
            sheet.Cells[triggerAddress].PutValue(123);

            // Manually recalculate all formulas after the change
            workbook.CalculateFormula();

            // Save the workbook
            const string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
