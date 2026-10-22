// Title: Create a C# Dictionary Mapping Excel Cell Addresses to Their Values by Enumerating Aspose.Cells Worksheet Cells
// AI Prompts: Generate C# code that opens an .xlsx workbook with Aspose.Cells, loops through worksheet.Cells, and adds each cell's Name and Value to a Dictionary<string, object>. | Provide a .NET console example that reads all cells from the first worksheet using Aspose.Cells and returns a dictionary where keys are cell addresses like "B2" and values are the corresponding cell contents.
// Common Searches: aspocells c# enumerate all cells and get address-value pairs | how to build dictionary of Excel cell names and values using Aspose.Cells | C# Aspose.Cells read worksheet cells into key value collection | extract cell address and content from .xlsx with Aspose.Cells in .NET | loop through Cells collection Aspose.Cells and store results in Dictionary
// Tags: Aspose.Cells enumerate worksheet cells | C# dictionary of Excel cell addresses | extract cell name and value Aspose.Cells | map cell address to content .NET | read all cells from .xlsx using Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// // Loads an Excel workbook with Aspose.Cells, iterates over every cell in the first worksheet, builds a Dictionary<string, object> that maps each cell's address (e.g., "A1") to its value, and prints the resulting key/value pairs.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Dictionary to map cell addresses (e.g., "A1") to their values
        Dictionary<string, object> cellAddressToValue = new Dictionary<string, object>();

        // Enumerate all cells in the worksheet's Cells collection
        foreach (Cell cell in worksheet.Cells)
        {
            // Store the cell's address and its value in the dictionary
            cellAddressToValue[cell.Name] = cell.Value;
        }

        // Example usage: print the dictionary contents
        foreach (var kvp in cellAddressToValue)
        {
            Console.WriteLine($"{kvp.Key}: {kvp.Value}");
        }
    }
}
