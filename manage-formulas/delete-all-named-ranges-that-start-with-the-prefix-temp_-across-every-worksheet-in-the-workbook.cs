// Title: C# – Remove all named ranges prefixed with “Temp_” using Aspose.Cells
// Description: Loads an Excel workbook, scans its NameCollection for defined names that start with the prefix Temp_ (case‑insensitive), deletes them in a single batch call, and saves the cleaned file.
// Keywords: Aspose.Cells | C# | .NET | delete named ranges | Temp_ prefix | NameCollection.Remove | Excel cleanup | remove temporary names
// Common Searches: Aspose.Cells delete named ranges with prefix | C# remove Temp_ names from Excel workbook | batch delete defined names Aspose.Cells | how to clean up temporary named ranges in .NET
// Developer Intent: Remove every defined name whose identifier begins with "Temp_" from the workbook.
// Use Cases: Sanitize generated reports by stripping internal temporary names before distribution. | Automate workbook cleanup in CI/CD pipelines to prevent leftover helper names. | Reduce file size and improve readability by deleting unused Temp_ ranges after calculations.
// AI Prompts: Write C# code that uses Aspose.Cells to delete all defined names starting with a given prefix in a workbook. | Explain the performance benefits of using NameCollection.Remove(string[]) versus removing names one by one. | Show how to filter NameCollection with LINQ to identify names that match a pattern and then delete them.

using Aspose.Cells;
using System;
using System.Collections.Generic;

// Loads an Excel workbook, scans its NameCollection for defined names that start with the prefix Temp_ (case‑insensitive), deletes them in a single batch call, and saves the cleaned file.
class DeleteTempNamedRanges
{
    static void Main()
    {
        // Load the workbook (adjust the file path as needed)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of all defined names in the workbook
        NameCollection names = workbook.Worksheets.Names;

        // Gather names that start with the prefix "Temp_"
        List<string> namesToDelete = new List<string>();
        foreach (Name name in names)
        {
            if (name.Text.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
            {
                namesToDelete.Add(name.Text);
            }
        }

        // Remove the collected names using the Remove(string[]) method
        if (namesToDelete.Count > 0)
        {
            names.Remove(namesToDelete.ToArray());
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
