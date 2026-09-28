// Title: Disable Calculation, Bulk‑Rename All Defined Names with a Prefix, and Recalculate Formulas in an Excel Workbook Using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to suspend automatic calculation, loop through the workbook's NameCollection, create new defined names prefixed with a given string, copy each name's RefersTo, visibility, and comment, delete the original name, then resume calculation and recalculate all formulas. | Show how to safely add a prefix to every named range in an XLSX file with Aspose.Cells while preventing intermediate formula evaluation, and then force a full formula recalculation after the renaming.
// Common Searches: aspocells c# how to rename all named ranges with a prefix and recalculate formulas | disable automatic calculation while updating defined names in Aspose.Cells | bulk rename Excel defined names using Aspose.Cells .NET and recalc workbook | add prefix to every name in workbook and force formula recalculation Aspose.Cells | C# Aspose.Cells rename NameCollection items without triggering calculation
// Tags: suspend calculation Aspose.Cells | bulk rename defined names .NET | add prefix to named ranges Aspose.Cells | recalculate formulas after name changes | iterate NameCollection backwards C#

using System;
using System.IO;
using Aspose.Cells;

// The program loads an XLSX workbook with Aspose.Cells, suspends automatic calculation, iterates the NameCollection in reverse, creates new names prefixed with a specified string while copying RefersTo, visibility, and comment, removes the original names, then calls CalculateFormula to recompute all formulas before saving the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prefix to add to each defined name
            const string prefix = "New_";

            // Get the collection of defined names
            NameCollection names = workbook.Worksheets.Names;

            // Rename each defined name by creating a new one with the prefix
            // and removing the original.
            for (int i = names.Count - 1; i >= 0; i--)
            {
                try
                {
                    // Existing name object
                    Aspose.Cells.Name oldName = names[i];
                    // The actual name string is stored in the Text property
                    string newNameStr = prefix + oldName.Text;

                    // Add a new name and obtain its index
                    int newIndex = names.Add(newNameStr);
                    Aspose.Cells.Name newName = names[newIndex];

                    // Copy properties from the old name
                    newName.RefersTo = oldName.RefersTo;
                    newName.IsVisible = oldName.IsVisible;
                    newName.Comment = oldName.Comment;

                    // Remove the old name
                    names.RemoveAt(i);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to rename a defined name at index {i}: {ex.Message}");
                }
            }

            // Recalculate all formulas
            workbook.CalculateFormula();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
