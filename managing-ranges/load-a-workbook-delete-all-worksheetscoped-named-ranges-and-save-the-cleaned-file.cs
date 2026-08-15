// Title: Remove all worksheet‑scoped named ranges from an Excel file using Aspose.Cells for .NET
// Description: Loads an existing workbook, extracts every worksheet‑level name from the Worksheets.Names collection, deletes them in a single operation, and saves the cleaned file. Ideal for preparing Excel files for distribution or further processing.
// Keywords: Aspose.Cells delete worksheet names | remove worksheet scoped named ranges .NET | clear Excel named ranges programmatically | Workbook.Worksheets.Names.Remove | clean Excel workbook Aspose.Cells | C# delete Excel named ranges
// Common Searches: how to delete all worksheet scoped names with Aspose.Cells | Aspose.Cells remove named ranges from workbook C# | clear worksheet level names in Excel using .NET | bulk delete Excel named ranges Aspose.Cells
// Developer Intent: Programmatically purge every worksheet‑level named range from a loaded workbook and write the sanitized version to disk.
// Use Cases: Strip internal names before sharing a template with external users. | Eliminate temporary ranges generated during data‑analysis pipelines. | Prepare a workbook for import into systems that do not support custom names.
// AI Prompts: Generate a C# method that accepts an input path, removes all worksheet‑scoped named ranges with Aspose.Cells, and returns the path of the cleaned file. | Show error‑handling best practices when deleting named ranges from a large Excel workbook using Aspose.Cells. | Explain how to verify that no worksheet‑level names remain after calling Worksheets.Names.Remove.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an existing workbook, extracts every worksheet‑level name from the Worksheets.Names collection, deletes them in a single operation, and saves the cleaned file. Ideal for preparing Excel files for distribution or further processing.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of worksheet‑scoped named ranges
        NameCollection worksheetNames = workbook.Worksheets.Names;

        // Collect the texts of all defined names
        List<string> namesToDelete = new List<string>();
        foreach (Name name in worksheetNames)
        {
            namesToDelete.Add(name.Text);
        }

        // Remove all collected names in one call
        worksheetNames.Remove(namesToDelete.ToArray());

        // Save the cleaned workbook
        workbook.Save("cleaned.xlsx");
    }
}
