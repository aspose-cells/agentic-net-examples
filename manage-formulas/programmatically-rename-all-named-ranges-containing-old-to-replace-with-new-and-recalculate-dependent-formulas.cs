// Title: Rename Excel named ranges containing 'Old' to 'New' and recalculate formulas using Aspose.Cells for .NET
// AI Prompts: Find all Name objects whose Text includes 'Old', replace the substring with 'New', and then call Workbook.CalculateFormula() to update dependent cells in C# with Aspose.Cells. | Batch rename named ranges by substituting a specific keyword and refresh all formulas in an Excel workbook using the Aspose.Cells .NET API.
// Common Searches: C# Aspose.Cells rename named ranges that contain a specific word | How to update named range names and recalculate formulas in an Excel file using Aspose.Cells | Batch change part of named range identifiers and refresh formulas with Aspose.Cells for .NET
// Tags: rename named ranges Aspose.Cells C# | replace substring in Name objects Aspose.Cells | recalculate workbook formulas Aspose.Cells | batch update Excel named ranges .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, identifies all named ranges whose name includes 'Old', replaces that part with 'New', recalculates all formulas to reflect the changes, and saves the modified file.
class RenameNamedRanges
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

            // Collect names that need to be renamed (avoid modifying collection while iterating)
            List<Name> namesToRename = new List<Name>();
            foreach (Name name in workbook.Worksheets.Names)
            {
                if (name.Text.Contains("Old"))
                {
                    namesToRename.Add(name);
                }
            }

            // Rename each collected name
            foreach (Name name in namesToRename)
            {
                string newNameText = name.Text.Replace("Old", "New");
                name.Text = newNameText; // Directly update the name
                // Comment is preserved automatically; no additional action required
            }

            // Recalculate formulas to reflect renamed ranges
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
