// Title: C# – Create a case‑insensitive dictionary that maps worksheet names to SheetId (TabId) using Aspose.Cells
// Description: This example shows how to generate a Workbook, assign custom TabId values to several worksheets, and populate a `Dictionary<string,int>` where each worksheet's name serves as the key and its SheetId (TabId) as the value. The resulting map enables O(1) retrieval of a worksheet's identifier by name, useful for fast look‑ups in large workbooks.
// Keywords: Aspose.Cells C# worksheet TabId | dictionary worksheet name to SheetId | fast worksheet ID lookup .NET | case‑insensitive worksheet map | Aspose.Cells API SheetId retrieval
// Common Searches: Aspose.Cells get worksheet TabId by name | C# dictionary for worksheet SheetId | lookup worksheet ID in Aspose.Cells workbook | case insensitive worksheet name to ID mapping | quick SheetId retrieval Aspose.Cells
// Developer Intent: Create a lookup table that returns a worksheet’s SheetId (TabId) when supplied with its name.
// Use Cases: Pass a worksheet’s SheetId to other Aspose.Cells methods that require a TabId. | Verify the existence of a worksheet before updating it, using O(1) dictionary checks. | Synchronize worksheet identifiers across multiple workbooks by storing a name‑to‑ID map.
// AI Prompts: Write C# code that builds a case‑insensitive dictionary mapping worksheet names to their TabId in an Aspose.Cells workbook. | Demonstrate how to safely retrieve a SheetId from the dictionary and handle missing worksheet names. | Show how to serialize the name‑to‑SheetId dictionary to JSON and reload it for later fast lookups.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsWorksheetIdLookup
{
    // This example shows how to generate a Workbook, assign custom TabId values to several worksheets, and populate a `Dictionary<string,int>` where each worksheet's name serves as the key and its SheetId (TabId) as the value. The resulting map enables O(1) retrieval of a worksheet's identifier by name, useful for fast look‑ups in large workbooks.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add a few worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "FirstSheet";
            workbook.Worksheets[0].TabId = 101; // example SheetId

            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.TabId = 202;

            Worksheet sheet3 = workbook.Worksheets.Add("ThirdSheet");
            sheet3.TabId = 303;

            // Dictionary to hold worksheet name -> SheetId (TabId) mapping
            Dictionary<string, int> sheetIdLookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // Populate the dictionary
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // ws.Name is the key, ws.TabId is the SheetId value
                sheetIdLookup[ws.Name] = ws.TabId;
            }

            // Example usage: retrieve SheetId by worksheet name
            string lookupName = "SecondSheet";
            if (sheetIdLookup.TryGetValue(lookupName, out int sheetId))
            {
                Console.WriteLine($"Worksheet \"{lookupName}\" has SheetId (TabId): {sheetId}");
            }
            else
            {
                Console.WriteLine($"Worksheet \"{lookupName}\" not found in the lookup dictionary.");
            }

            // Save the workbook (optional)
            workbook.Save("WorksheetIdLookupDemo.xlsx");
        }
    }
}
