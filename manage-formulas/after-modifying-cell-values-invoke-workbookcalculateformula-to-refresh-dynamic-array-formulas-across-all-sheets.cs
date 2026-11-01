// Title: Recalculate all formulas, including dynamic arrays, after modifying cells with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# recalculate workbook after changing cell values | How to update dynamic array formulas in every worksheet using Aspose.Cells | C# example of refreshing Excel formulas across multiple sheets with Aspose.Cells | Programmatically force formula refresh after cell edit in Aspose.Cells .NET | Recompute all workbook formulas after modifying cells in Aspose.Cells C#
// Tags: recalculate workbook formulas Aspose.Cells | dynamic array formula refresh .NET | update all worksheet formulas C# | cell value modification Aspose.Cells | Excel workbook formula refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook or creates a new one, changes the values of cells A1 and B2 on the first worksheet, calls workbook.CalculateFormula to recompute every formula—including dynamic array formulas—across all sheets, and then saves the workbook, handling loading and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException ex)
                {
                    // Handle password‑protected or corrupted files
                    Console.WriteLine($"Error loading workbook: {ex.Message}");
                    return;
                }
            }
            else
            {
                // Create a new workbook with a default worksheet
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Modify some cell values
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(123);
            sheet.Cells["B2"].PutValue("Sample");

            // Recalculate all formulas (including dynamic array formulas) across all sheets
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
