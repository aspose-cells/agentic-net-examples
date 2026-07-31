// Title: C# – Rename Excel named ranges that start with "Temp" by adding an "Archive_" prefix using Aspose.Cells
// Description: Load a workbook, loop through its NameCollection, find names whose Text begins with "Temp", prepend "Archive_" to each, optionally sort the collection, and save the updated file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells rename named range | C# Excel batch rename | prefix Temp Archive | NameCollection loop | modify defined names | sort renamed names | Excel automation .NET
// Common Searches: Aspose.Cells rename named ranges with specific prefix | C# loop through NameCollection and change name.Text | Add Archive_ prefix to Temp named ranges in Excel | How to sort renamed defined names using Aspose.Cells | Batch rename Excel named ranges programmatically
// Developer Intent: Rename every named range that begins with "Temp" by prefixing it with "Archive_".
// Use Cases: Archive temporary calculation ranges before publishing a workbook | Migrate legacy named ranges to a new naming convention in bulk | Standardize temporary data names after automated processing
// AI Prompts: Write C# code with Aspose.Cells that finds all defined names starting with "Temp", prepends "Archive_" to each, sorts the NameCollection, and saves the workbook. | Explain how to filter a NameCollection for a given prefix and rename those entries without affecting other names in Aspose.Cells. | Create a reusable C# method that accepts a workbook path, a source prefix, and a target prefix to rename matching named ranges using Aspose.Cells.

using Aspose.Cells;
using System;
using System.Collections.Generic;

// Load a workbook, loop through its NameCollection, find names whose Text begins with "Temp", prepend "Archive_" to each, optionally sort the collection, and save the updated file with Aspose.Cells for .NET.
class RenameTempNames
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of defined names in the workbook
        NameCollection names = workbook.Worksheets.Names;

        // Collect the names that start with "Temp"
        List<Name> namesToRename = new List<Name>();
        foreach (Name name in names)
        {
            if (!string.IsNullOrEmpty(name.Text) && name.Text.StartsWith("Temp"))
            {
                namesToRename.Add(name);
            }
        }

        // Rename each collected name by prefixing "Archive_"
        foreach (Name name in namesToRename)
        {
            name.Text = "Archive_" + name.Text;
        }

        // Optional: sort the names after renaming
        names.Sort();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
