using System;
using System.Collections.Generic;
using Aspose.Cells;

class ValidateNamedRanges
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of all defined names in the workbook
        NameCollection names = workbook.Worksheets.Names;

        // Dictionary to collect occurrences of each name (case‑insensitive)
        Dictionary<string, List<Name>> nameMap = new Dictionary<string, List<Name>>(StringComparer.OrdinalIgnoreCase);

        // Populate the dictionary
        foreach (Name name in names)
        {
            string key = name.Text; // The name text
            if (!nameMap.ContainsKey(key))
                nameMap[key] = new List<Name>();
            nameMap[key].Add(name);
        }

        bool duplicatesFound = false;

        // Scan for entries that appear more than once
        foreach (var entry in nameMap)
        {
            if (entry.Value.Count > 1)
            {
                duplicatesFound = true;
                Console.WriteLine($"Duplicate name: '{entry.Key}' appears {entry.Value.Count} times.");
                // Show each duplicate's reference
                foreach (Name dup in entry.Value)
                {
                    Console.WriteLine($"  RefersTo: {dup.RefersTo}");
                }
            }
        }

        if (!duplicatesFound)
        {
            Console.WriteLine("No duplicate named ranges found.");
        }

        // Save the workbook (unchanged) if needed
        workbook.Save("output.xlsx");
    }
}