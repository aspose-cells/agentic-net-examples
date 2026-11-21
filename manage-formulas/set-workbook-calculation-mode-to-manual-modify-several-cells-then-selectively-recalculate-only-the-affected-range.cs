// Title: How to set manual calculation mode and recalculate a specific range in an Aspose.Cells workbook using C#
// Common Searches: Aspose.Cells C# set workbook calculation mode to manual | Recalculate only a range of formulas after changing cells with Aspose.Cells | Worksheet.CalculateFormula example for specific cells in .NET | Partial formula evaluation to improve performance in Aspose.Cells | Avoid full workbook recalculation when updating values in C#
// Tags: manual calculation mode Aspose.Cells | Worksheet.CalculateFormula range | partial formula recalculation C# | update cell values Aspose.Cells | selective formula evaluation .NET | performance optimization Excel workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, sets values in A1 and B1, defines formulas in C1 and D1, switches to manual calculation mode, and uses Worksheet.CalculateFormula to recalculate only the C1:D1 range before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Modify several cells
            sheet.Cells["A1"].PutValue(10);               // Input value
            sheet.Cells["B1"].PutValue(5);                // Input value
            sheet.Cells["C1"].Formula = "=A1+B1";         // Formula dependent on A1 and B1
            sheet.Cells["D1"].Formula = "=C1*2";          // Formula dependent on C1

            // Selectively recalculate only the affected range (C1:D1)
            sheet.CalculateFormula("C1:D1");

            // Determine output path and ensure its directory exists
            string outputPath = "ManualCalculationResult.xlsx";
            string? outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
