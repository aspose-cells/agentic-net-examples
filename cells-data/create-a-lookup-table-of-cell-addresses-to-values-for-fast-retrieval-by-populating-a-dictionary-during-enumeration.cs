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
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("John");
            cells["B1"].PutValue(28);
            cells["A2"].PutValue(DateTime.Now);
            cells["B2"].PutValue(3.1415);
            cells["C3"].PutValue(true);

            // Dictionary to hold address -> value mapping
            Dictionary<string, object> lookup = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            // Get enumerator for all cells in the worksheet (enumeration rule)
            IEnumerator enumerator = cells.GetEnumerator();
            while (enumerator.MoveNext())
            {
                // Cast the current object to Cell
                Cell cell = (Cell)enumerator.Current;

                // Only add cells that actually contain a value
                if (cell.Value != null)
                {
                    // Use cell.Name (e.g., "A1") as the key and cell.Value as the value
                    lookup[cell.Name] = cell.Value;
                }
            }

            // Demonstrate fast retrieval from the dictionary
            Console.WriteLine("Lookup results:");
            foreach (var kvp in lookup)
            {
                Console.WriteLine($"{kvp.Key} => {kvp.Value}");
            }

            // Example of retrieving a specific cell value by address
            string addressToFind = "B1";
            if (lookup.TryGetValue(addressToFind, out object foundValue))
            {
                Console.WriteLine($"\nValue at {addressToFind}: {foundValue}");
            }
            else
            {
                Console.WriteLine($"\nAddress {addressToFind} not found in the lookup table.");
            }

            // Save the workbook (save rule)
            workbook.Save("LookupDemo.xlsx");
        }
    }
}