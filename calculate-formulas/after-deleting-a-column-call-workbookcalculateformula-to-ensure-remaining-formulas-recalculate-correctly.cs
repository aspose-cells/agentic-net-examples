// Title: Delete a column from an Excel worksheet and recalculate formulas using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells delete column and force formula recalculation | How to use Workbook.CalculateFormula after removing a worksheet column | Remove column B from Excel file with Aspose.Cells and update dependent formulas | Aspose.Cells recalculate all formulas after structural changes in .NET | Delete worksheet column and refresh formulas using Aspose.Cells C# example
// Tags: remove worksheet column Aspose.Cells C# | Workbook.CalculateFormula after structural modification | full formula recalculation .NET Excel | update dependent formulas after column removal | exception handling for missing or protected workbook Aspose.Cells | save modified workbook after column removal C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, removes column B from the first worksheet, forces a complete formula recalculation with Workbook.CalculateFormula, and saves the updated workbook while handling file‑not‑found, password‑protected, and other Aspose.Cells exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook (no password supplied). If the file is password‑protected,
            // an exception will be caught below.
            Workbook workbook = new Workbook(inputPath);

            // Delete column B (zero‑based index 1) from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.DeleteColumn(1);

            // Recalculate all formulas after the column deletion
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., password protection)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
