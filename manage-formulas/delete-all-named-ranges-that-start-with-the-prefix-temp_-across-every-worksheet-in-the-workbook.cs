using Aspose.Cells;
using System;
using System.Collections.Generic;

class DeleteTempNamedRanges
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of all defined names in the workbook
        NameCollection names = workbook.Worksheets.Names;

        // Gather the names that start with the prefix "Temp_"
        List<string> namesToDelete = new List<string>();
        foreach (Name name in names)
        {
            // Name.Text holds the name string
            if (name.Text.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
            {
                namesToDelete.Add(name.Text);
            }
        }

        // Remove the collected names in a single call
        if (namesToDelete.Count > 0)
        {
            names.Remove(namesToDelete.ToArray());
        }

        // Save the workbook after removal
        workbook.Save("output.xlsx");
    }
}