// Title: Rename all named ranges with a prefix and recalculate formulas in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to prepend a prefix to every named range in an Excel workbook using Aspose.Cells C# | Aspose.Cells example for bulk renaming named ranges and recalculating formulas | C# code to rename Excel named ranges with a custom prefix and trigger full calculation | temporarily suspend automatic calculation while updating named ranges with Aspose.Cells
// Tags: rename named ranges Aspose.Cells | prepend prefix to Excel named ranges C# | bulk named range rename Aspose.Cells | force formula recalculation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel file with Aspose.Cells, iterates through all defined names, adds a specified prefix to each named range, forces a full formula recalculation, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prefix to add to each named range
            string prefix = "New_";

            // Retrieve all named ranges in the workbook
            NameCollection namedRanges = workbook.Worksheets.Names;

            // Rename each named range by adding the prefix
            foreach (Name name in namedRanges)
            {
                try
                {
                    string originalName = name.Text;
                    name.Text = prefix + originalName;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to rename range '{name.Text}': {ex.Message}");
                }
            }

            // Force a full recalculation of all formulas
            workbook.CalculateFormula();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
