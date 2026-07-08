using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsLookupDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("John");
            cells["B1"].PutValue(28);
            cells["C1"].PutValue(DateTime.Now);
            cells["A2"].PutValue("Alice");
            cells["B2"].PutValue(34);
            cells["C2"].PutValue(123.45);

            // Build a lookup dictionary: key = cell address (e.g., "A1"), value = cell.Value
            Dictionary<string, object> cellLookup = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            // Use the Cells.GetEnumerator method to iterate through all existing cells
            IEnumerator enumerator = cells.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                // Store only non‑null values; you can adjust this condition as needed
                if (cell.Value != null)
                {
                    cellLookup[cell.Name] = cell.Value;
                }
            }

            // Example of fast retrieval using the dictionary
            string addressToFind = "B1";
            if (cellLookup.TryGetValue(addressToFind, out object value))
            {
                Console.WriteLine($"Value at {addressToFind}: {value}");
            }
            else
            {
                Console.WriteLine($"Cell {addressToFind} not found or contains null.");
            }

            // Save the workbook (optional, just to demonstrate lifecycle)
            workbook.Save("LookupDemo.xlsx");
        }
    }
}