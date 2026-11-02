// Title: Recalculate all formulas after deleting a column with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete a column and refresh remaining formulas | How to force formula recalculation after column removal in .NET | Workbook.CalculateFormula usage after worksheet column deletion | Update Excel formulas after deleting column B using Aspose.Cells | C# Aspose.Cells recalculate formulas after modifying worksheet structure
// Tags: delete column Aspose.Cells C# | Workbook.CalculateFormula .NET | formula recalculation after column deletion | Excel worksheet structure change Aspose.Cells | refresh formulas after column removal

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook (creating a simple one if missing), deletes column B from the first worksheet, calls Workbook.CalculateFormula to recalculate all remaining formulas, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var tempWs = tempWb.Worksheets[0];
                tempWs.Cells["A1"].PutValue("Sample Data");
                tempWs.Cells["B1"].PutValue("Will be deleted");
                tempWb.Save(inputPath);
            }

            // Load the workbook from the file (no password required).
            var workbook = new Workbook(inputPath);

            // Delete column B (zero‑based index 1) from the first worksheet.
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells.DeleteColumn(1);

            // Recalculate all formulas after the column deletion.
            workbook.CalculateFormula();

            // Save the updated workbook.
            workbook.Save(outputPath);
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
