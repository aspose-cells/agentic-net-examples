// Title: Create a fast cell address‑to‑value lookup dictionary in C# with Aspose.Cells
// Description: This example builds a case‑insensitive Dictionary<string, object> by enumerating all non‑empty cells in a worksheet, storing each cell's address (e.g., "A1") and its value for instant retrieval, then displays the map and saves the workbook.
// Keywords: Aspose.Cells lookup dictionary | C# cell address mapping | enumerate worksheet cells .NET | fast cell value retrieval | case‑insensitive cell lookup | dictionary of Excel cells
// Common Searches: how to map Excel cell addresses to values in C# | Aspose.Cells enumerate cells and store in dictionary | quick lookup of cell values by address .NET | create cell address lookup table using Aspose.Cells | case insensitive cell address dictionary C#
// Developer Intent: Generate a dictionary that maps each populated cell’s address to its value for rapid, address‑based access.
// Use Cases: Cache worksheet data for high‑performance read‑only calculations. | Implement custom formulas that require instant lookup of cell values by address. | Validate input by checking the existence of specific cell addresses in a pre‑built map.
// AI Prompts: Write C# code that creates a case‑insensitive Dictionary<string, object> from an Aspose.Cells worksheet, ignoring empty cells. | Extend the lookup to include cell style properties (font, background color) alongside the value. | Provide a method that receives a cell address string and returns the stored value with proper error handling.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

// This example builds a case‑insensitive Dictionary<string, object> by enumerating all non‑empty cells in a worksheet, storing each cell's address (e.g., "A1") and its value for instant retrieval, then displays the map and saves the workbook.
class LookupTableExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some sample data
        cells["A1"].PutValue("Hello");
        cells["B1"].PutValue(123);
        cells["A2"].PutValue(DateTime.Now);
        cells["C3"].PutValue(3.14);

        // Dictionary to store cell address -> cell value
        Dictionary<string, object> lookup = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        // Enumerate all cells using the provided GetEnumerator method
        IEnumerator enumerator = cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;
            if (cell.Value != null)
            {
                // Use the cell's Name (e.g., "A1") as the key
                lookup[cell.Name] = cell.Value;
            }
        }

        // Demonstrate fast retrieval from the dictionary
        Console.WriteLine("Lookup Table Contents:");
        foreach (KeyValuePair<string, object> entry in lookup)
        {
            Console.WriteLine($"{entry.Key} => {entry.Value}");
        }

        // Optional: save the workbook to verify the data
        workbook.Save("LookupTableDemo.xlsx");
    }
}
