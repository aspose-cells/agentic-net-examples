// Title: Recalculate Aspose.Cells workbook formulas only after a specific cell value changes using C#
// Common Searches: Aspose.Cells calculate formulas only when cell A1 changes C# | C# how to trigger Workbook.CalculateFormula after a particular cell update with Aspose.Cells | prevent automatic formula recalculation in Aspose.Cells until manual call in .NET | conditional formula evaluation based on cell change Aspose.Cells C# example | event‑free way to recalculate dependent formulas after setting a cell value in Aspose.Cells
// Tags: manual formula calculation Aspose.Cells | cell‑specific recalculation C# | Workbook.CalculateFormula after cell update | conditional workbook recalculation .NET | dependent formula evaluation Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, sets a formula in B1 that depends on A1, modifies an unrelated cell without triggering calculation, then updates A1 and calls Workbook.CalculateFormula() to recalculate dependent formulas before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a formula that depends on cell A1
            sheet.Cells["B1"].Formula = "=A1*2";

            // Change a different cell (no recalculation will happen)
            sheet.Cells["C1"].PutValue(10); // No trigger

            // Change the specific cell A1 and recalculate formulas manually
            sheet.Cells["A1"].PutValue(5);
            workbook.CalculateFormula();

            // Save the workbook to a file
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
