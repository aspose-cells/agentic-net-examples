// Title: Aspose.Cells for .NET – Delete all worksheet‑scoped (or all) named ranges from an Excel workbook
// Description: Loads a workbook, extracts every defined name from the Worksheets.Names collection, removes them with NameCollection.Remove, and saves the cleaned file. Includes optional filtering for worksheet‑scoped names only.
// Keywords: Aspose.Cells delete named ranges | remove worksheet scoped names .NET | clear defined names C# | Workbook.Worksheets.Names.Remove | Aspose.Cells clean workbook | C# Excel named range removal
// Common Searches: how to delete worksheet scoped named ranges using Aspose.Cells | remove all defined names from an Excel file C# | Aspose.Cells clear named ranges before saving | C# code to purge named ranges in a workbook | Aspose.Cells delete names collection
// Developer Intent: Eliminate every defined name—optionally only worksheet‑scoped—from a workbook and write the result to a new file.
// Use Cases: Strip legacy named ranges from a template before distribution. | Prevent reference errors after data migration by clearing obsolete names. | Prepare an Excel file for systems that do not support named ranges.
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, removes only worksheet‑scoped named ranges, and saves the file. | Show how to use NameCollection.Remove(string[]) to delete all defined names in a workbook. | Explain how to check the IsWorksheetScoped property before removing a name.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, extracts every defined name from the Worksheets.Names collection, removes them with NameCollection.Remove, and saves the cleaned file. Includes optional filtering for worksheet‑scoped names only.
    public class DeleteWorksheetScopedNames
    {
        public static void Run()
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file \"{inputFile}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Get the collection of defined names (includes worksheet‑scoped names)
                NameCollection names = workbook.Worksheets.Names;

                // Collect the texts of all names to be removed
                List<string> namesToRemove = new List<string>();
                foreach (Name name in names)
                {
                    // If you need to filter only worksheet‑scoped names, uncomment the next line
                    // if (name.IsWorksheetScoped) namesToRemove.Add(name.Text);
                    
                    // For removing all defined names, add every name
                    namesToRemove.Add(name.Text);
                }

                // Remove the collected names using the Remove(string[]) method
                if (namesToRemove.Count > 0)
                {
                    names.Remove(namesToRemove.ToArray());
                }

                // Save the cleaned workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public static class Program
    {
        public static void Main(string[] args)
        {
            DeleteWorksheetScopedNames.Run();
        }
    }
}
