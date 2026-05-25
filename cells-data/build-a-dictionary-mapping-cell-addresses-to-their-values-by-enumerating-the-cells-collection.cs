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

        // Add some sample data to the worksheet
        cells["A1"].PutValue("Hello");
        cells["B1"].PutValue(123);
        cells["C2"].PutValue(DateTime.Now);
        cells["D4"].PutValue(true);

        // Dictionary to hold cell address -> value mappings
        Dictionary<string, object> cellValues = new Dictionary<string, object>();

        // Enumerate all instantiated cells in the worksheet
        IEnumerator enumerator = cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;
            // Store only cells that have a non‑null value
            if (cell.Value != null)
            {
                cellValues[cell.Name] = cell.Value;
            }
        }

        // Display the collected address‑value pairs
        foreach (var kvp in cellValues)
        {
            Console.WriteLine($"{kvp.Key} = {kvp.Value}");
        }

        // Save the workbook (optional, demonstrates the required save rule)
        workbook.Save("CellDictionaryDemo.xlsx");
    }
}