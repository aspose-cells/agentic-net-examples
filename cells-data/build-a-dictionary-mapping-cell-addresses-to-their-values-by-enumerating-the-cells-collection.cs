// Title: Create a C# Dictionary of Excel Cell Addresses and Values by Enumerating Aspose.Cells Cells
// Description: Loads a workbook, accesses the first worksheet, enumerates instantiated cells with Cells.GetEnumerator(), and adds each non‑empty cell’s address (Name) and its Value to a Dictionary<string,object>. The map is printed and the workbook saved.
// Keywords: Aspose.Cells | C# enumerate cells | cell address dictionary | Excel cell values | Cells.GetEnumerator | populate dictionary | worksheet cell mapping | Aspose.Cells API | C# Excel automation
// Common Searches: Aspose.Cells enumerate cells dictionary C# | how to get cell address and value Aspose.Cells | C# map Excel cells to dictionary | retrieve non‑empty cells Aspose.Cells | list populated cells with addresses using Aspose.Cells
// Developer Intent: Build a key‑value map of populated cell addresses to their values in a worksheet using Aspose.Cells for .NET.
// Use Cases: Serialize worksheet data to JSON for web APIs. | Compare two sheets by creating address‑value maps and detecting differences. | Generate an audit report that lists every filled cell with its location and content. | Feed cell data into a custom validation engine that expects a dictionary input.
// AI Prompts: Write C# code with Aspose.Cells that enumerates all cells in the first worksheet and stores each non‑null cell’s address and value in a Dictionary<string,object>. | Show how to filter the enumeration so only cells containing data are added to the dictionary, and then output the address‑value pairs. | Explain how to modify the example to store the formatted text (displayed string) instead of the raw Value for each cell in the dictionary.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

// Loads a workbook, accesses the first worksheet, enumerates instantiated cells with Cells.GetEnumerator(), and adds each non‑empty cell’s address (Name) and its Value to a Dictionary<string,object>. The map is printed and the workbook saved.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Dictionary to map cell addresses to their values
        Dictionary<string, object> cellValues = new Dictionary<string, object>();

        // Enumerate all instantiated cells in the worksheet
        IEnumerator enumerator = cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;
            // Store only cells that contain a value
            if (cell.Value != null)
            {
                cellValues[cell.Name] = cell.Value;
            }
        }

        // Example: display the collected address‑value pairs
        foreach (KeyValuePair<string, object> kvp in cellValues)
        {
            Console.WriteLine($"{kvp.Key}: {kvp.Value}");
        }

        // Save the workbook (if any changes were made)
        workbook.Save("output.xlsx");
    }
}
