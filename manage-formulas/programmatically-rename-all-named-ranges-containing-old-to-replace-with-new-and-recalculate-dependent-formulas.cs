// Title: Rename Excel named ranges containing 'Old' to 'New' and recalculate dependent formulas using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# rename all named ranges that include a specific substring | How to replace part of a named range name and keep formula references in .NET | Bulk rename Excel named ranges with Aspose.Cells and ensure unique names | Recalculate formulas after renaming named ranges using Aspose.Cells C# | Programmatically update named ranges and recalculate dependent formulas in an Excel workbook
// Tags: rename named ranges Aspose.Cells | replace substring in named range C# | recalculate formulas after renaming Aspose.Cells | ensure unique named range names Aspose.Cells | load and save workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook, finds all named ranges whose names contain "Old", renames each to replace "Old" with "New" while guaranteeing unique names, recalculates all formulas that depend on those ranges, and saves the modified workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook (provide password via LoadOptions if required)
            LoadOptions loadOptions = new LoadOptions();
            // loadOptions.Password = "yourPassword"; // uncomment and set if needed
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Collect named ranges containing "Old"
            List<Name> namesToRename = new List<Name>();
            foreach (Name name in workbook.Worksheets.Names)
            {
                if (!string.IsNullOrEmpty(name.Text) && name.Text.Contains("Old"))
                {
                    namesToRename.Add(name);
                }
            }

            // Rename each collected named range, ensuring uniqueness
            foreach (Name name in namesToRename)
            {
                string newName = name.Text.Replace("Old", "New");

                if (workbook.Worksheets.Names[newName] == null)
                {
                    name.Text = newName;
                }
                else
                {
                    int suffix = 1;
                    string uniqueName = newName;
                    while (workbook.Worksheets.Names[uniqueName] != null)
                    {
                        uniqueName = $"{newName}_{suffix}";
                        suffix++;
                    }
                    name.Text = uniqueName;
                }
            }

            // Recalculate formulas that depend on renamed ranges
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
