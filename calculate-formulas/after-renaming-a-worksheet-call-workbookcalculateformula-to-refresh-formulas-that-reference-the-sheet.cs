// Title: Rename an Excel worksheet and recalculate formulas with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells rename worksheet and update formulas | How to refresh formula references after sheet rename using Aspose.Cells | Workbook.CalculateFormula after changing worksheet name in .NET | Aspose.Cells recalculate all formulas when worksheet name changes | Example code to rename sheet and recalc formulas in Aspose.Cells C#
// Tags: Aspose.Cells rename worksheet | Workbook.CalculateFormula method usage | refresh formula references C# | Excel sheet rename formula update | recalculate formulas Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file or creates a new workbook, ensures a worksheet exists, renames the first worksheet to "RenamedSheet", calls Workbook.CalculateFormula to refresh any formulas that reference the renamed sheet, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load workbook if the file exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException ex)
                {
                    // Handle cases such as password‑protected files
                    Console.WriteLine($"Failed to load '{inputPath}': {ex.Message}");
                    Console.WriteLine("Creating a new workbook instead.");
                    workbook = new Workbook();
                }
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                workbook.Worksheets.Add();
            }

            // Rename the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "RenamedSheet";

            // Recalculate all formulas to refresh references
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
