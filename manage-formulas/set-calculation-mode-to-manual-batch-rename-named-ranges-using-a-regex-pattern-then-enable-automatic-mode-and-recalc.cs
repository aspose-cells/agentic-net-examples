// Title: Batch rename Excel named ranges with a regular expression and recalculate formulas using Aspose.Cells for .NET
// Common Searches: how to rename Excel named ranges using regex with Aspose.Cells C# | Aspose.Cells batch rename named ranges and recalculate workbook | C# replace named range prefix Old_ to New_ in an Excel file using Aspose.Cells | programmatically update multiple named ranges and recalculate formulas Aspose.Cells | save workbook after renaming named ranges with Aspose.Cells .NET
// Tags: Aspose.Cells rename named ranges regex | C# batch rename Excel named ranges | Aspose.Cells recalculate formulas after rename | Excel named range prefix replacement Aspose.Cells | Aspose.Cells save workbook after modifications

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an Excel workbook, scans all defined names, and uses a regular expression to rename any named range that starts with "Old_" to a new name prefixed with "New_" while preserving the original RefersTo reference. After applying the renames, it triggers a full formula recalculation and saves the updated workbook to a new file, handling common I/O and processing errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // -----------------------------------------------------------------
        // 1. Batch rename named ranges using a regular expression
        //    Example: rename ranges that start with "Old_" to "New_..."
        // -----------------------------------------------------------------
        string pattern = @"^Old_(.*)$";          // regex pattern to match old names
        string replacement = "New_$1";           // replacement pattern

        // Collect the names that need to be renamed to avoid modifying the collection while iterating
        var renameMap = new List<(string OldName, string NewName, string RefersTo)>();

        foreach (Name name in workbook.Worksheets.Names)
        {
            string newName = Regex.Replace(name.Text, pattern, replacement);
            if (!newName.Equals(name.Text, StringComparison.Ordinal))
            {
                renameMap.Add((name.Text, newName, name.RefersTo));
            }
        }

        // Apply the renames
        foreach (var item in renameMap)
        {
            try
            {
                // Add the new name (using the overload that only takes the name)
                int index = workbook.Worksheets.Names.Add(item.NewName);
                // Set the same reference as the old name
                workbook.Worksheets.Names[index].RefersTo = item.RefersTo;
                // Remove the old name
                workbook.Worksheets.Names.Remove(item.OldName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to rename '{item.OldName}' to '{item.NewName}': {ex.Message}");
            }
        }

        // -----------------------------------------------------------------
        // 2. Recalculate the workbook after renaming
        // -----------------------------------------------------------------
        try
        {
            workbook.CalculateFormula();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during workbook recalculation: {ex.Message}");
        }

        // -----------------------------------------------------------------
        // 3. Save the modified workbook
        // -----------------------------------------------------------------
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
