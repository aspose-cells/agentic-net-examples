// Title: Create a C# dictionary that maps Excel cell addresses to values while enumerating an Aspose.Cells worksheet
// Common Searches: c# Aspose.Cells build dictionary of cell address to value | enumerate worksheet cells with Aspose.Cells and create fast lookup | retrieve Excel cell value by address using Aspose.Cells dictionary | skip empty cells when iterating Aspose.Cells for performance | how to map non‑empty cells to a lookup table in Aspose.Cells C#
// Tags: Aspose.Cells cell address to value map | C# worksheet cell iteration Aspose.Cells | filter empty cells Aspose.Cells | high‑performance cell value lookup C# | dictionary<string,object> for Excel cells Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

namespace Example
{
    // The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet, iterates over all cells, and builds a Dictionary<string, object> that maps each non‑empty cell's address (e.g., "A1") to its value, then saves the workbook.
    public class CellLookupBuilder
    {
        public void Build()
        {
            try
            {
                string inputPath = "input.xlsx";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Dictionary to store address‑to‑value mappings
                Dictionary<string, object> lookup = new Dictionary<string, object>();

                // Enumerate all cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Skip empty cells for faster processing
                    if (cell.Value != null)
                    {
                        // Use the cell's address (e.g., "A1") as the key
                        lookup[cell.Name] = cell.Value;
                    }
                }

                // Example of retrieving a value by address
                // if (lookup.TryGetValue("B2", out object value))
                //     Console.WriteLine($"B2 = {value}");

                // Save workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            new CellLookupBuilder().Build();
        }
    }
}
