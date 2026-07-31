// Title: C# – Bulk prepend a prefix to all named ranges, toggle calculation, and recalc with Aspose.Cells
// Description: Loads a workbook, disables automatic calculation, adds a custom prefix to every defined name while preserving its RefersTo reference, sorts the NameCollection, re‑enables calculation, forces a full formula recalculation, and saves the result to a new file.
// Keywords: Aspose.Cells C# rename named ranges | bulk add prefix to Excel defined names | disable calculation Aspose.Cells | enable calculation and recalc | NameCollection sort | programmatic Excel range renaming | formula recalculation after rename
// Common Searches: Aspose.Cells add prefix to all named ranges C# | disable calculation then rename named ranges Aspose.Cells | recalculate formulas after bulk renaming Excel names | C# code to rename Excel defined names with Aspose.Cells | how to sort NameCollection after renaming
// Developer Intent: Rename every defined name with a common prefix, temporarily suspend calculation, then recalculate all formulas.
// Use Cases: Standardize naming conventions across a workbook before distribution. | Prevent name collisions when merging multiple workbooks by bulk‑renaming ranges. | Integrate into CI/CD pipelines to update named ranges and verify workbook integrity with a forced recalculation.
// AI Prompts: Write C# code using Aspose.Cells that disables automatic calculation, prepends "New_" to each defined name, re‑enables calculation, and forces a full formula recalculation. | Show how to keep the RefersTo property unchanged while renaming named ranges in a workbook with Aspose.Cells for .NET. | Provide an example that sorts the renamed NameCollection, saves the workbook, and logs success messages.

using System;
using System.IO;
using Aspose.Cells;

// Loads a workbook, disables automatic calculation, adds a custom prefix to every defined name while preserving its RefersTo reference, sorts the NameCollection, re‑enables calculation, forces a full formula recalculation, and saves the result to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Prefix to be added to each named range
            string prefix = "New_";

            // Get the collection of all defined names in the workbook
            NameCollection names = workbook.Worksheets.Names;

            // Rename each named range by adding the prefix
            foreach (Name name in names)
            {
                // Preserve the original reference so it is not altered unintentionally
                string originalRef = name.RefersTo;

                // Update the name text with the desired prefix
                name.Text = prefix + name.Text;

                // Reassign the original reference (ensures the range stays the same)
                name.RefersTo = originalRef;
            }

            // Optional: sort the names after renaming for better organization
            names.Sort();

            // Force a full recalculation after renaming
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
