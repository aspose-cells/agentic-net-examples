// Title: Export All Worksheet Cells to JSON Using Cells Enumerator in Aspose.Cells for .NET (C#)
// Description: The sample creates a workbook, adds sample data, obtains an IEnumerator from Worksheet.Cells, iterates each Cell, calls Cell.ToJson to get the cell’s address and value as JSON, combines the results into a JSON array, and prints the final string to the console.
// Keywords: Aspose.Cells | C# | Cells enumerator | Cell.ToJson | Excel to JSON | export worksheet to JSON | iterate cells .NET | JSON array of cells
// Common Searches: Aspose.Cells export cells to JSON | How to use Cells.GetEnumerator in C# | Cell.ToJson example Aspose.Cells | Convert Excel worksheet to JSON with Aspose | Generate JSON array of cell addresses Aspose.Cells
// Developer Intent: Create a JSON array that lists every cell’s address and value by enumerating the Cells collection of a worksheet.
// Use Cases: Send worksheet data as a JSON payload in a REST API response. | Log all cell contents for debugging or audit trails during workbook processing. | Provide front‑end JavaScript applications with a ready‑to‑use JSON representation of Excel data.
// AI Prompts: Show how to filter out empty cells before adding them to the JSON array. | Modify the code to write the JSON output to a file instead of the console. | Demonstrate deserializing the generated JSON back into a .NET dictionary of cell addresses and values.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, adds sample data, obtains an IEnumerator from Worksheet.Cells, iterates each Cell, calls Cell.ToJson to get the cell’s address and value as JSON, combines the results into a JSON array, and prints the final string to the console.
    public class CellsEnumeratorToJsonDemo
    {
        public static void Run()
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

            // Collect JSON representation of each cell
            List<string> cellJsonList = new List<string>();

            while (enumerator.MoveNext())
            {
                // Current item is a Cell
                Cell cell = (Cell)enumerator.Current;

                // Convert the cell to JSON using the built‑in ToJson method
                string cellJson = cell.ToJson();

                cellJsonList.Add(cellJson);
            }

            // Combine individual cell JSON objects into a JSON array
            string allCellsJson = "[" + string.Join(",", cellJsonList) + "]";

            // Output the resulting JSON
            Console.WriteLine("All cells JSON:");
            Console.WriteLine(allCellsJson);
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                CellsEnumeratorToJsonDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
