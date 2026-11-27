// Title: How to rename Excel named ranges that include "Old" to "New" and refresh dependent formulas using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells rename named range containing specific text | replace substring in Excel named range and recalculate formulas with Aspose.Cells | programmatically update named ranges and recalculate workbook formulas in .NET | avoid duplicate named range when renaming using Aspose.Cells | iterate named range collection backwards Aspose.Cells C#
// Tags: rename named ranges Aspose.Cells C# | replace substring in Excel named range | recalculate workbook formulas Aspose.Cells | prevent duplicate named range Aspose.Cells | iterate named range collection backward

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates the named‑range collection in reverse, renames any range whose name contains "Old" by replacing it with "New" (skipping the rename if the new name already exists), recalculates all formulas to reflect the changes, and saves the updated file.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of named ranges
            NameCollection namedRanges = workbook.Worksheets.Names;

            // Iterate backwards to safely rename items
            for (int i = namedRanges.Count - 1; i >= 0; i--)
            {
                // Retrieve the named range object
                Name namedRange = namedRanges[i];

                // Get the current name of the range
                string currentName = namedRange.Text;

                // Check if the name contains the substring "Old"
                if (!string.IsNullOrEmpty(currentName) && currentName.Contains("Old"))
                {
                    // Create the new name by replacing "Old" with "New"
                    string newName = currentName.Replace("Old", "New");

                    // Ensure the new name does not already exist to avoid duplicates
                    if (namedRanges[newName] == null)
                    {
                        // Rename the named range
                        namedRange.Text = newName;
                    }
                    // else: conflict – skip renaming
                }
            }

            // Recalculate all formulas that may depend on the renamed ranges
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
            // Catch any unexpected errors and display a message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
