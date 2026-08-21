// Title: C# – Remove Underscore‑Prefixed Defined Names from an Excel Workbook with Aspose.Cells
// Description: Loads an Excel file using Aspose.Cells, scans the workbook‑ and worksheet‑scoped NameCollection for any defined name that begins with "_", removes those names in a single operation, validates that none remain, and saves the cleaned workbook to a new file.
// Keywords: Aspose.Cells remove defined names | filter underscore named ranges C# | NameCollection Remove method | delete workbook scoped names Aspose | Excel named range underscore prefix | C# Aspose.Cells example | GitHub Aspose.Cells defined name filter
// Common Searches: how to delete underscore‑prefixed defined names with Aspose.Cells | C# filter out named ranges that start with '_' when loading a workbook | remove workbook and worksheet scoped names beginning with underscore | Aspose.Cells example for cleaning named ranges
// Developer Intent: Load a workbook, strip all defined names whose text starts with an underscore, confirm their removal, and save the updated file.
// Use Cases: Sanitize legacy templates by removing internal helper names before distribution. | Enforce naming conventions in exported Excel files by automatically deleting temporary underscore‑prefixed ranges. | Prepare workbook assets for publishing by cleaning up hidden or system‑generated named ranges.
// AI Prompts: Generate C# code using Aspose.Cells that loads a workbook and removes every defined name starting with '_' in one call. | Show how to use NameCollection.Remove(string[]) to filter out underscore‑prefixed names after loading a workbook. | Explain how to verify that no names beginning with '_' remain and how to handle validation errors in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDefinedNameFilter
{
    // Loads an Excel file using Aspose.Cells, scans the workbook‑ and worksheet‑scoped NameCollection for any defined name that begins with "_", removes those names in a single operation, validates that none remain, and saves the cleaned workbook to a new file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook (should contain defined names, some starting with "_")
                string sourcePath = "InputWithNames.xlsx";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook with default load options (all data including defined names)
                LoadOptions loadOptions = new LoadOptions();
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Access the collection of defined names (both workbook‑scoped and worksheet‑scoped)
                NameCollection names = workbook.Worksheets.Names;

                // Gather the texts of all names that start with '_' 
                List<string> namesToRemove = new List<string>();
                foreach (Name name in names)
                {
                    if (name.Text.StartsWith("_"))
                    {
                        namesToRemove.Add(name.Text);
                    }
                }

                // Remove the collected names in a single call using the provided Remove(string[]) method
                if (namesToRemove.Count > 0)
                {
                    names.Remove(namesToRemove.ToArray());
                }

                // Verify that no defined name beginning with '_' remains in the collection
                foreach (Name name in names)
                {
                    if (name.Text.StartsWith("_"))
                    {
                        throw new InvalidOperationException($"Name '{name.Text}' was not removed as expected.");
                    }
                }

                // Save the workbook to a new file to confirm the changes
                string resultPath = "OutputWithoutUnderscoreNames.xlsx";
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to: {resultPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
