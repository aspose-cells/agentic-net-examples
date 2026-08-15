// Title: Aspose.Cells C# – Validate Unique Named Ranges and Report Duplicates
// Description: Loads an Excel workbook with Aspose.Cells, enumerates all defined names, detects case‑insensitive duplicates, prints each repeated name with its index positions, and optionally removes them before saving the file.
// Keywords: Aspose.Cells duplicate named ranges C# | validate unique defined names .NET | list duplicate Excel names Aspose | detect repeated named ranges programmatically | remove duplicate defined names Aspose.Cells
// Common Searches: how to find duplicate named ranges using Aspose.Cells for .NET | C# code to check unique defined names in an Excel workbook | list indices of repeated named ranges Aspose.Cells | remove duplicate named ranges automatically
// Developer Intent: Identify and output any named ranges that share the same identifier within a workbook.
// Use Cases: Ensure template integrity by flagging duplicate named ranges before distribution. | Generate a diagnostic log for automated cleanup of Excel files. | Programmatically purge duplicate names and save a cleaned workbook.
// AI Prompts: Create C# code with Aspose.Cells that lists all defined names and highlights those occurring more than once, showing their indices. | Write a method that removes duplicate named ranges after reporting them and saves the workbook to a new file. | Explain how to switch the duplicate‑name detection from case‑insensitive to case‑sensitive in the provided example.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace NamedRangeValidator
{
    // Loads an Excel workbook with Aspose.Cells, enumerates all defined names, detects case‑insensitive duplicates, prints each repeated name with its index positions, and optionally removes them before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Get the collection of all defined names in the workbook
            NameCollection names = workbook.Worksheets.Names;

            // Dictionary to track name occurrences and their indices
            var nameOccurrences = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);

            // Iterate through all names and record their positions
            for (int i = 0; i < names.Count; i++)
            {
                Name name = names[i];
                string text = name.Text;

                if (!nameOccurrences.ContainsKey(text))
                {
                    nameOccurrences[text] = new List<int>();
                }
                nameOccurrences[text].Add(i);
            }

            // Flag to indicate if any duplicates were found
            bool duplicatesFound = false;

            // Report duplicate names
            foreach (var kvp in nameOccurrences)
            {
                if (kvp.Value.Count > 1)
                {
                    duplicatesFound = true;
                    Console.WriteLine($"Duplicate name \"{kvp.Key}\" found at indices: {string.Join(", ", kvp.Value)}");
                }
            }

            if (!duplicatesFound)
            {
                Console.WriteLine("All named ranges have unique names.");
            }

            // Optional: remove duplicates after reporting
            // names.RemoveDuplicateNames();

            // Save the workbook if any modifications were made (e.g., after removal)
            // workbook.Save("output.xlsx");
        }
    }
}
