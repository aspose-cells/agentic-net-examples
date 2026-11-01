// Title: Build a case‑insensitive lookup dictionary of Excel cell addresses to values with Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates over every Cell in an Aspose.Cells worksheet and populates a Dictionary<string, object> where the key is the cell's address (e.g., "A2") and the value is the cell's content. | Show how to retrieve a value from the dictionary by providing a cell address string, including proper handling when the address does not exist. | Include sample data creation, dictionary construction, a demonstration lookup, and code to save the workbook as an .xlsx file.
// Common Searches: asp.net c# create dictionary of cell addresses using Aspose.Cells | how to map Excel cell names to values with Aspose.Cells .NET | case insensitive lookup of Excel cell by address Aspose.Cells | enumerate all cells in Aspose.Cells worksheet and store in Dictionary | fast retrieve cell value by address Aspose.Cells C#
// Tags: Aspose.Cells enumerate cells to dictionary | case-insensitive cell address lookup Aspose.Cells | C# dictionary mapping Excel cell names | fast cell value retrieval Aspose.Cells | save workbook after building lookup Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsLookupExample
{
    // The example creates a Workbook, fills cells A1:B4 with sample data, iterates through every cell in the first worksheet, and stores each cell's address and value in a case‑insensitive Dictionary<string, object> for rapid lookup. It then demonstrates retrieving the value at address B3 and saves the workbook as LookupExample.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Initialize a new workbook (you could also load an existing one using Workbook(string fileName))
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data for demonstration purposes
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Score");
            worksheet.Cells["A2"].PutValue("Alice");
            worksheet.Cells["B2"].PutValue(85);
            worksheet.Cells["A3"].PutValue("Bob");
            worksheet.Cells["B3"].PutValue(92);
            worksheet.Cells["A4"].PutValue("Charlie");
            worksheet.Cells["B4"].PutValue(78);

            // Create a dictionary to hold cell address (e.g., "A2") as key and its value as value
            Dictionary<string, object> cellLookup = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            // Enumerate all cells in the worksheet and populate the dictionary
            foreach (Cell cell in worksheet.Cells)
            {
                // Skip empty cells if desired; here we include all cells
                cellLookup[cell.Name] = cell.Value;
            }

            // Example usage: fast retrieval of a cell's value by its address
            string addressToLookup = "B3";
            if (cellLookup.TryGetValue(addressToLookup, out object value))
            {
                Console.WriteLine($"Value at {addressToLookup}: {value}");
            }
            else
            {
                Console.WriteLine($"Cell {addressToLookup} not found in the lookup table.");
            }

            // Optionally, save the workbook to a file (using the standard save method)
            workbook.Save("LookupExample.xlsx");
        }
    }
}
