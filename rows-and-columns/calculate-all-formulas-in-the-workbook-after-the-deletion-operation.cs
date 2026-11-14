// Title: Delete a specific row and recalculate all formulas in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete row and recalculate workbook formulas | How to remove a row at index 5 and update formulas with Aspose.Cells | Recalculate all formulas after deleting rows in an .xlsx using Aspose.Cells | C# code to delete rows and run Workbook.CalculateFormula in Aspose.Cells | Example of row deletion followed by formula evaluation with Aspose.Cells
// Tags: delete rows with Aspose.Cells C# | Workbook.CalculateFormula method | recalculate formulas after row deletion | Excel worksheet row removal using Aspose.Cells | formula evaluation after worksheet changes

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, deletes the row at zero‑based index 5 on the first worksheet, invokes Workbook.CalculateFormula() to recompute all formulas, and then saves the updated file.
class Program
{
    static void Main(string[] args)
    {
        // Determine input and output file paths (use arguments if provided)
        string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
        string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Delete 1 row starting at index 5 (zero‑based)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.DeleteRows(5, 1);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook processed successfully. Saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
