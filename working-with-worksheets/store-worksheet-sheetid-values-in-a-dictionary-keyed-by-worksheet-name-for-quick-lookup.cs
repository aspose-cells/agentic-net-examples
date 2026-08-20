// Title: C# – Build a case‑insensitive dictionary of worksheet names to SheetId (TabId) with Aspose.Cells
// Description: This example creates a Workbook, assigns custom TabId values to three worksheets, and populates a case‑insensitive Dictionary<string,int> where each key is the worksheet name and each value is the corresponding SheetId. It demonstrates fast lookup, prints the mapping, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | Worksheet TabId | SheetId lookup | case‑insensitive dictionary | fast worksheet ID retrieval | sample code | GitHub example
// Common Searches: Aspose.Cells map worksheet name to TabId | C# dictionary worksheet name SheetId | quick lookup worksheet ID Aspose.Cells | case insensitive worksheet ID dictionary .NET | how to get worksheet TabId without looping
// Developer Intent: Create a dictionary that maps each worksheet’s name to its TabId (SheetId) for instant retrieval in C#.
// Use Cases: Retrieve a worksheet’s TabId in O(1) time for conditional processing. | Synchronize workbook identifiers with external systems that reference worksheets by name. | Validate worksheet existence before applying formatting, formulas, or data updates.
// AI Prompts: Generate C# code using Aspose.Cells that builds a case‑insensitive dictionary of worksheet names to TabId values and shows how to query it. | Show how to keep the name‑to‑TabId dictionary up‑to‑date when worksheets are added or removed at runtime in an Aspose.Cells workbook. | Provide a GitHub‑style README snippet explaining the purpose and performance benefits of the dictionary lookup.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// This example creates a Workbook, assigns custom TabId values to three worksheets, and populates a case‑insensitive Dictionary<string,int> where each key is the worksheet name and each value is the corresponding SheetId. It demonstrates fast lookup, prints the mapping, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample worksheets and assign TabId values (SheetId)
        workbook.Worksheets[0].Name = "First";
        workbook.Worksheets[0].TabId = 101;

        Worksheet sheet2 = workbook.Worksheets.Add("Second");
        sheet2.TabId = 202;

        Worksheet sheet3 = workbook.Worksheets.Add("Third");
        sheet3.TabId = 303;

        // Build a dictionary keyed by worksheet name with TabId as the value
        Dictionary<string, int> sheetIdLookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (Worksheet ws in workbook.Worksheets)
        {
            sheetIdLookup[ws.Name] = ws.TabId;
        }

        // Demonstrate quick lookup
        Console.WriteLine("Worksheet TabId lookup:");
        foreach (var kvp in sheetIdLookup)
        {
            Console.WriteLine($"Name: {kvp.Key}, TabId: {kvp.Value}");
        }

        // Save the workbook (optional)
        workbook.Save("SheetIdLookup.xlsx");
    }
}
