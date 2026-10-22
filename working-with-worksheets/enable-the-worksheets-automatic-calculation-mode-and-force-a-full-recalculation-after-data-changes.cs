// Title: Enable automatic calculation mode and trigger a full workbook recalculation with Aspose.Cells in C#
// AI Prompts: Generate C# code that sets Workbook.Settings.CalcMode to Automatic and then calls Workbook.CalculateFormula to recompute all formulas after updating cells using Aspose.Cells. | Show me how to turn on auto‑calculate for an Aspose.Cells workbook and force a complete formula recompute in .NET. | Write a C# snippet that enables automatic calculation for a worksheet, modifies a cell value, and then runs a full recalculation with Aspose.Cells.
// Common Searches: Aspose.Cells C# enable auto calculate and recalculate workbook after cell update | Set CalcMode to Automatic and force full formula recalculation in Aspose.Cells .NET | How to trigger CalculateFormula after changing data in an Aspose.Cells workbook | Automatic calculation mode Aspose.Cells C# example
// Tags: Workbook.Settings.CalcMode Automatic Aspose.Cells | full formula recompute Aspose.Cells .NET | auto‑calculate setup C# Aspose.Cells | force CalculateFormula execution .NET | Aspose.Cells calculation mode configuration example

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), accesses the first worksheet, forces a complete formula recalculation with Workbook.CalculateFormula (which respects the workbook's calculation settings), and saves the updated file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Load an existing template if it exists; otherwise create a new workbook
            string templatePath = "Template.xlsx";
            Workbook workbook = File.Exists(templatePath) ? new Workbook(templatePath) : new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
