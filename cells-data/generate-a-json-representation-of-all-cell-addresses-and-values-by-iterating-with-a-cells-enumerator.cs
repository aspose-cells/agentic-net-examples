// Title: Export Non‑Empty Cell Addresses and Values to JSON Using a Cells Enumerator in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills sample data, obtains an IEnumerator for the worksheet's Cells collection, iterates through each Cell, skips empty entries, captures the cell address (Name) and its Value, and serializes the result to a pretty‑printed JSON string with System.Text.Json.
// Keywords: Aspose.Cells | C# | .NET | cells enumerator | JSON serialization | cell address | cell value | export worksheet to JSON | System.Text.Json | iterate worksheet cells
// Common Searches: Aspose.Cells enumerate cells to JSON C# | How to get cell address and value with Aspose.Cells | Export worksheet data as JSON using Aspose.Cells | Iterate over Cells collection with IEnumerator | Skip empty cells when converting Aspose.Cells to JSON
// Developer Intent: Generate a JSON array that lists the addresses and values of all populated cells in a worksheet by using a Cells enumerator.
// Use Cases: Provide worksheet data to web APIs or front‑end grids in JSON format. | Create a lightweight snapshot of spreadsheet contents for logging or debugging. | Transfer non‑empty cell information between services without exporting the whole file.
// AI Prompts: Write C# code that uses Aspose.Cells to iterate over a worksheet's Cells collection with an enumerator and outputs a JSON array of address/value pairs, ignoring empty cells. | Show how to extend the example to include each cell's formula and data type in the JSON output. | Explain how to stream the generated JSON directly to a file instead of printing it to the console in the Aspose.Cells enumerator sample.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills sample data, obtains an IEnumerator for the worksheet's Cells collection, iterates through each Cell, skips empty entries, captures the cell address (Name) and its Value, and serializes the result to a pretty‑printed JSON string with System.Text.Json.
    public class CellsEnumeratorToJsonDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Age");
                cells["A2"].PutValue("John");
                cells["B2"].PutValue(30);
                cells["A3"].PutValue("Alice");
                cells["B3"].PutValue(25);

                // Get the enumerator for the Cells collection
                IEnumerator enumerator = cells.GetEnumerator();

                // Collect address-value pairs for non‑empty cells
                var cellList = new List<Dictionary<string, object>>();

                while (enumerator.MoveNext())
                {
                    Cell cell = (Cell)enumerator.Current;

                    // Skip cells without a value
                    if (cell.Value == null) continue;

                    var entry = new Dictionary<string, object>
                    {
                        { "Address", cell.Name },
                        { "Value", cell.Value }
                    };
                    cellList.Add(entry);
                }

                // Convert the list to a formatted JSON string
                string jsonResult = JsonSerializer.Serialize(cellList, new JsonSerializerOptions { WriteIndented = true });

                // Output the JSON representation
                Console.WriteLine(jsonResult);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            CellsEnumeratorToJsonDemo.Run();
        }
    }
}
