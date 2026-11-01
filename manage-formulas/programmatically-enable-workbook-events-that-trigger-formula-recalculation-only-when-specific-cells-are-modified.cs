// Title: Recalculate workbook formulas only after modifying specific cells using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# recalculate formulas only when certain cells are changed | trigger workbook.CalculateFormula after updating specific cells in .NET | conditional formula recalculation based on cell edits with Aspose.Cells | C# code to recalculate Excel formulas after modifying cell A1 using Aspose.Cells | avoid unnecessary formula calculation in Aspose.Cells workbook
// Tags: conditional formula recalculation Aspose.Cells | workbook.CalculateFormula C# | cell update trigger formula Aspose.Cells | selective formula evaluation .NET | Excel workbook event handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates an Excel workbook, modifies cells A1 and C3 on the first worksheet, then explicitly calls workbook.CalculateFormula() to recalculate all formulas only after those specific cell updates, and finally saves the workbook to output.xlsx, all wrapped in exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Load an existing workbook if it exists; otherwise create a new one
            string inputPath = "input.xlsx";
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example modifications
            sheet.Cells["A1"].PutValue(10); // Change that would normally trigger recalculation
            sheet.Cells["C3"].PutValue(5);  // Change that does not affect formulas

            // Manually recalculate all formulas
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
