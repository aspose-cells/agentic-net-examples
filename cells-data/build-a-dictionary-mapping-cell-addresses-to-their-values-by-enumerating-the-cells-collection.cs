using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some sample data
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Age");
        cells["A2"].PutValue("John");
        cells["B2"].PutValue(30);
        cells["A3"].PutValue("Alice");
        cells["B3"].PutValue(25);

        // Dictionary to map cell addresses to their values
        Dictionary<string, object> addressValueMap = new Dictionary<string, object>();

        // Enumerate the Cells collection using Cells.GetEnumerator()
        IEnumerator enumerator = cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;
            // Add only cells that contain a non‑null value
            if (cell.Value != null)
            {
                addressValueMap[cell.Name] = cell.Value;
            }
        }

        // Output the dictionary contents
        foreach (KeyValuePair<string, object> kvp in addressValueMap)
        {
            Console.WriteLine($"{kvp.Key}: {kvp.Value}");
        }

        // Save the workbook (optional, just to demonstrate lifecycle usage)
        workbook.Save("CellAddressValueMap.xlsx");
    }
}