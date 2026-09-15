// Title: Check for duplicate named ranges in an Excel workbook with Aspose.Cells for .NET and report their RefersTo addresses
// AI Prompts: Write C# code using Aspose.Cells that iterates all workbook named ranges, identifies names that appear more than once (ignoring case), and prints each duplicate with its RefersTo address. | Create a reusable C# method that receives a Workbook object and returns a dictionary where keys are duplicate named‑range names and values are lists of their RefersTo formulas.
// Common Searches: Aspose.Cells C# find duplicate named range definitions in an .xlsx file | how to list named ranges and detect name collisions with Aspose.Cells .NET | C# program to validate that named range names are unique in an Excel workbook using Aspose.Cells
// Tags: detect duplicate named ranges Aspose.Cells | validate unique named range names C# | case-insensitive named range check Aspose.Cells | list workbook named ranges .NET

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, groups all named ranges case‑insensitively, and prints any names that occur multiple times together with each RefersTo address.
class NamedRangeValidator
{
    static void Main(string[] args)
    {
        // Path to the workbook to validate
        string filePath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(filePath);

        // Dictionary to track occurrences of each named range name (case‑insensitive)
        Dictionary<string, List<string>> nameOccurrences = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        // Iterate through all named ranges in the workbook
        foreach (Name namedRange in workbook.Worksheets.Names)
        {
            string rangeName = namedRange.Text;          // The name of the range
            string refersTo = namedRange.RefersTo;       // The address the name refers to

            // Record the occurrence
            if (!nameOccurrences.ContainsKey(rangeName))
            {
                nameOccurrences[rangeName] = new List<string> { refersTo };
            }
            else
            {
                nameOccurrences[rangeName].Add(refersTo);
            }
        }

        // Report any duplicate names
        bool duplicatesFound = false;
        foreach (var entry in nameOccurrences)
        {
            if (entry.Value.Count > 1)
            {
                duplicatesFound = true;
                Console.WriteLine($"Duplicate named range \"{entry.Key}\" found {entry.Value.Count} times:");
                for (int i = 0; i < entry.Value.Count; i++)
                {
                    Console.WriteLine($"  Instance {i + 1}: RefersTo = {entry.Value[i]}");
                }
            }
        }

        if (!duplicatesFound)
        {
            Console.WriteLine("All named ranges have unique names.");
        }
    }
}
