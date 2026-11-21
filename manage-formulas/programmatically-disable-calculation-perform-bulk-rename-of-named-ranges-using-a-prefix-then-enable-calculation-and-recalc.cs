// Title: Add a prefix to all named ranges, ensure unique names, and recalculate formulas with Aspose.Cells for .NET
// Common Searches: how to prepend a string to every named range in an Excel workbook using Aspose.Cells C# | bulk rename defined names in an .xlsx file with Aspose.Cells without breaking formulas | disable automatic calculation while updating named ranges with Aspose.Cells .NET | recalculate all formulas after renaming named ranges in Aspose.Cells C# | ensure unique named range names when adding a prefix using Aspose.Cells
// Tags: add prefix to named ranges Aspose.Cells | bulk rename defined names .xlsx C# | turn off calculation mode Aspose.Cells | recalculate workbook formulas Aspose.Cells | ensure unique named range names C# | save modified workbook Aspose.Cells

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The example loads an existing workbook, iterates through its NameCollection, prepends a specified prefix to each defined name while guaranteeing uniqueness, recalculates all formulas, and saves the updated workbook. It also demonstrates safe collection iteration and handling of missing input files and output directories.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prefix to add to each named range
            const string prefix = "New_";

            // Get the collection of defined names
            NameCollection names = workbook.Worksheets.Names;

            // Iterate over a snapshot of the collection size to avoid issues while renaming
            int originalCount = names.Count;
            for (int i = 0; i < originalCount; i++)
            {
                Name name = names[i];
                string originalName = name.Text;
                string newName = prefix + originalName;

                // Ensure the new name does not conflict with existing names
                int suffix = 1;
                while (names.Any(n => string.Equals(n.Text, newName, StringComparison.OrdinalIgnoreCase)))
                {
                    newName = $"{prefix}{originalName}_{suffix}";
                    suffix++;
                }

                // Rename the defined name
                name.Text = newName;
            }

            // Recalculate all formulas in the workbook
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
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
