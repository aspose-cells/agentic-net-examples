// Title: Recalculate all formulas in an Excel workbook using Aspose.Cells Workbook.CalculateFormula (C#)
// Common Searches: C# Aspose.Cells calculate all formulas in a workbook without options | How to recalculate Excel formulas after loading a file with Aspose.Cells .NET | Default formula evaluation example using Workbook.CalculateFormula in Aspose.Cells | Run full workbook calculation after modifying cells with Aspose.Cells C# | Aspose.Cells recalculate formulas and save the workbook programmatically
// Tags: Workbook.CalculateFormula default calculation | Aspose.Cells formula recalculation C# | Excel workbook full calculation Aspose | load and save workbook Aspose.Cells | C# evaluate Excel formulas Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file (or creates a new workbook), calls Workbook.CalculateFormula() to evaluate every formula using Aspose.Cells' default calculation engine, and then saves the workbook, handling CellsException, IOException, and generic errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if the file exists; otherwise create a new workbook.
            if (File.Exists(inputPath))
            {
                // If the workbook is password protected, provide the password via LoadOptions.
                // Here we assume no password; adjust LoadOptions.Password if needed.
                var loadOptions = new LoadOptions();
                workbook = new Workbook(inputPath, loadOptions);
            }
            else
            {
                // Create a new workbook with a default worksheet.
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Calculate all formulas using default calculation settings.
            workbook.CalculateFormula();

            // Save the workbook after calculation.
            workbook.Save(outputPath);
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (IOException ex)
        {
            Console.WriteLine($"IO error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
