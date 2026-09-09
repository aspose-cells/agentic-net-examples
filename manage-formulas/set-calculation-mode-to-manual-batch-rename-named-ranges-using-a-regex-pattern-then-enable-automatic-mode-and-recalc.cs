// Title: Set workbook calculation to Manual, batch rename Excel named ranges with a regex, then restore Automatic mode and recalculate formulas using Aspose.Cells for .NET
// AI Prompts: Configure the Aspose.Cells workbook to use Manual calculation, iterate through all defined names, apply a case‑insensitive regex that changes names starting with 'Old_' to 'New_', ensure each new name is unique, then switch the calculation mode back to Automatic and trigger a full formula recalculation. | Modify the sample code to temporarily disable automatic formula evaluation, perform bulk regex‑based renaming of named ranges, resolve naming collisions by appending numeric suffixes, re‑enable automatic calculation, and call CalculateFormula to update all dependent cells.
// Common Searches: how to turn off automatic formula calculation in Aspose.Cells before editing named ranges | C# Aspose.Cells rename multiple named ranges using regular expression | batch update Excel defined names and resolve duplicate name errors with Aspose.Cells | recalculate workbook formulas after changing named ranges in Aspose.Cells .NET | set calculation mode to manual then back to automatic in Aspose.Cells example
// Tags: disable automatic formula evaluation Aspose.Cells | regex based bulk rename of named ranges C# | handle named range naming collisions Aspose.Cells | enable automatic calculation after batch update Aspose.Cells | recalculate all formulas programmatically Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, temporarily switches formula calculation to Manual, iterates over all defined names, uses a case‑insensitive regex to rename ranges matching 'Old_*' to 'New_*' while ensuring uniqueness, restores Automatic calculation mode, forces a full formula recalculation, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Regex pattern to rename named ranges (e.g., Old_Something -> New_Something)
            Regex renamePattern = new Regex(@"Old_(.*)", RegexOptions.IgnoreCase);

            // Iterate through all named ranges in the workbook
            foreach (Name namedRange in workbook.Worksheets.Names)
            {
                string currentName = namedRange.Text;

                // Check if the name matches the pattern
                if (renamePattern.IsMatch(currentName))
                {
                    // Generate the new name
                    string newName = renamePattern.Replace(currentName, "New_$1");

                    // Ensure the new name does not already exist
                    if (workbook.Worksheets.Names[newName] == null)
                    {
                        namedRange.Text = newName;
                    }
                    else
                    {
                        // Resolve conflict by appending a numeric suffix
                        int suffix = 1;
                        string uniqueName = newName;
                        while (workbook.Worksheets.Names[uniqueName] != null)
                        {
                            uniqueName = $"{newName}_{suffix}";
                            suffix++;
                        }
                        namedRange.Text = uniqueName;
                    }
                }
            }

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

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
