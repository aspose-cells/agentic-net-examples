// Title: Create a C# Dictionary of Excel Cell Addresses and Their Values by Enumerating Aspose.Cells Worksheet Cells
// AI Prompts: Generate C# code that iterates over worksheet.Cells with Aspose.Cells, captures each cell's Name and non‑null Value, and stores them in a Dictionary<string, object>. | Show how to build a lookup table of cell addresses to values from an Aspose.Cells workbook for subsequent processing.
// Common Searches: how to iterate over all populated cells in Aspose.Cells and collect their addresses and values in C# | C# Aspose.Cells example for converting worksheet data to a Dictionary<string, object> | retrieve non‑empty cell names and values from an Excel file using Aspose.Cells enumeration
// Tags: Aspose.Cells enumerate worksheet cells to dictionary | C# extract cell address and value Aspose.Cells | dictionary mapping Excel cell names to values using Aspose | non‑null cell value extraction Aspose.Cells C# | store worksheet data in Dictionary<string, object>

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDictionaryExample
{
    // The sample creates a workbook, optionally adds data, then uses worksheet.Cells.GetEnumerator() to loop through all instantiated cells. For each cell with a non‑null value, it adds the cell's Name and Value to a Dictionary<string, object>. The dictionary is printed and the workbook can be saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data – this part can be omitted if loading an existing file
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Age");
            worksheet.Cells["A2"].PutValue("Alice");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["A3"].PutValue("Bob");
            worksheet.Cells["B3"].PutValue(25);

            // Dictionary to hold cell address (e.g., "A1") and its value
            Dictionary<string, object> cellValues = new Dictionary<string, object>();

            // Enumerate through all instantiated cells in the worksheet
            IEnumerator enumerator = worksheet.Cells.GetEnumerator(); // GetEnumerator
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                // Store the cell name and its value (null values are ignored)
                if (cell.Value != null)
                {
                    cellValues[cell.Name] = cell.Value;
                }
            }

            // Example usage: print the dictionary contents
            foreach (KeyValuePair<string, object> kvp in cellValues)
            {
                Console.WriteLine($"{kvp.Key}: {kvp.Value}");
            }

            // Save the workbook (optional, demonstrates the save rule)
            workbook.Save("DictionaryExample.xlsx"); // save
        }
    }
}
