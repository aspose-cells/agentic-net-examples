// Title: Rename a worksheet and recalculate formulas with Aspose.Cells for .NET (C#)
// Common Searches: how to rename a worksheet and refresh formulas using Aspose.Cells in C# | Aspose.Cells Workbook.CalculateFormula after changing sheet name | update formula references after sheet rename Aspose.Cells C# | recalculate all formulas when worksheet name is changed in Aspose.Cells | C# Aspose.Cells rename worksheet then recalc workbook formulas
// Tags: Worksheet rename Aspose.Cells C# | Workbook.CalculateFormula after worksheet rename | refresh formula references Aspose.Cells | recalculate all formulas Aspose.Cells workbook | save workbook after renaming sheet Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file (or creates a new workbook), renames the first worksheet, calls Workbook.CalculateFormula to update any formulas that reference the renamed sheet, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a workbook with a default worksheet
            }

            // Rename the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "RenamedSheet";

            // Recalculate all formulas to reflect the renamed sheet
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (CellsException ex)
        {
            Console.Error.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
