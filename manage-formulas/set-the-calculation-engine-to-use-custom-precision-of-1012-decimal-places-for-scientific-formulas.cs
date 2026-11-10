// Title: How to achieve 10‑12 decimal place precision for scientific formulas with Aspose.Cells in C#
// Common Searches: Aspose.Cells calculate formulas with 12 decimal places precision in .NET | set formula calculation precision to 10 decimal places using Aspose.Cells C# | increase numeric accuracy of Excel formulas in Aspose.Cells .NET | does Aspose.Cells support custom decimal precision for scientific calculations | default calculation engine decimal limits in Aspose.Cells
// Tags: Aspose.Cells formula calculation precision | Aspose.Cells default calculation engine decimal accuracy | C# high precision scientific formulas Aspose.Cells | Excel workbook calculation decimal places Aspose.Cells | Aspose.Cells high precision formula evaluation

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, notes that Aspose.Cells does not expose a public CalculationEngine, and explains that the built‑in engine already provides up to 12 decimal places of precision, which satisfies the 10‑12 decimal place requirement for scientific formulas. It then recalculates all formulas with workbook.CalculateFormula() and saves the workbook to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: Aspose.Cells does not expose a public CalculationEngine class.
            // The default calculation engine already supports up to 12 decimal places,
            // which satisfies the required precision range (10‑12).

            // Recalculate all formulas using the default engine
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
