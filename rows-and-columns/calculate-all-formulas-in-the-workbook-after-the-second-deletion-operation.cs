// Title: Recalculate all formulas in an Excel workbook after deleting a specific row and column with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# recalculate workbook formulas after DeleteRows | How to update formulas after removing a column using Aspose.Cells .NET | Calculate all formulas in Excel file after structural changes with Aspose.Cells | C# code to delete a row and a column then refresh formulas in a workbook | Aspose.Cells DeleteColumns followed by Workbook.CalculateFormula example
// Tags: Aspose.Cells delete rows recalculate formulas | Workbook.CalculateFormula after DeleteColumns | C# Excel row column removal update formulas | Aspose.Cells structural changes formula recalculation | Excel workbook formula refresh Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, deletes the third row (zero‑based index 2) and the fourth column (zero‑based index 3) from the first worksheet, calls Workbook.CalculateFormula() to recompute all formulas, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Delete 1 row starting at index 2 (zero‑based)
                workbook.Worksheets[0].Cells.DeleteRows(2, 1);

                // Delete 1 column starting at index 3 (zero‑based)
                // The DeleteColumns method requires a third parameter in recent versions
                workbook.Worksheets[0].Cells.DeleteColumns(3, 1, false);

                // Recalculate all formulas after deletions
                workbook.CalculateFormula();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed successfully. Saved to '{outputPath}'.");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"File error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
