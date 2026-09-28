// Title: Generate a JSON array of populated cell addresses and values by enumerating Aspose.Cells worksheet cells in C#
// AI Prompts: Write C# code that uses Aspose.Cells to iterate through all non‑empty cells in a worksheet and output each cell's address and value as a JSON object. | Extend the enumeration to include each cell's formula and number format, adding those details to the JSON representation. | Create a reusable C# method that returns a JSON string containing an array of all populated cells from a given Aspose.Cells worksheet.
// Common Searches: how to export non empty cells from an Aspose.Cells worksheet to JSON in C# | Aspose.Cells C# enumerate cells and get address and value as JSON | C# convert Excel worksheet cells to JSON array using Aspose.Cells enumerator | retrieve cell formulas and formatting with Aspose.Cells and serialize to JSON
// Tags: Aspose.Cells cell enumeration JSON conversion | C# export worksheet data as JSON with Aspose.Cells | Cell.ToJson for populated cells Aspose.Cells | non‑empty cell address extraction Aspose.Cells C# | serialize Excel cells to JSON using Aspose.Cells

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, populating sample cells, using the Cells enumerator to walk through non‑empty cells, converting each cell to JSON with Cell.ToJson, and assembling the results into a JSON array string.
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

                // Populate some sample data
                cells["A1"].PutValue(100);
                cells["B2"].PutValue("Hello");
                cells["C3"].PutValue(DateTime.Now);
                cells["D4"].PutValue(3.14);

                // Get the enumerator for the Cells collection
                IEnumerator enumerator = cells.GetEnumerator();

                // Collect JSON representation of each non‑empty cell
                List<string> cellJsonList = new List<string>();
                while (enumerator.MoveNext())
                {
                    Cell cell = (Cell)enumerator.Current;
                    if (cell.Value != null)
                    {
                        // Cell.ToJson returns a JSON string for the individual cell
                        cellJsonList.Add(cell.ToJson());
                    }
                }

                // Combine individual cell JSON strings into a JSON array
                string jsonResult = "[" + string.Join(",", cellJsonList) + "]";

                // Output the final JSON
                Console.WriteLine(jsonResult);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CellsEnumeratorToJsonDemo.Run();
        }
    }
}
