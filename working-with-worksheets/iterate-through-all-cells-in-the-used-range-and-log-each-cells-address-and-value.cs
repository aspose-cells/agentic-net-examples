// Title: C# – Enumerate all cells in a worksheet’s used range with Aspose.Cells and log address/value
// Description: Demonstrates how to create a workbook, fill sample data, obtain an enumerator for the used range via Cells.GetEnumerator(), iterate each Cell, output its address (Name) and value to the console, and finally save the file.
// Keywords: Aspose.Cells enumerate used range C# | C# iterate worksheet cells Aspose | log cell address Aspose.Cells | Aspose.Cells GetEnumerator example | save workbook after cell iteration | Aspose.Cells console output cells | C# read non‑empty cells Aspose
// Common Searches: how to loop through used range cells Aspose.Cells .NET | Aspose.Cells get cell address while iterating | C# enumerate all populated cells in Aspose workbook | log each cell value Aspose.Cells console | save workbook after iterating cells Aspose
// Developer Intent: Iterate over every populated cell in a worksheet’s used range and output its address and value.
// Use Cases: Debugging: quickly view all non‑empty cells and their contents. | Auditing: generate a log of data entries before further processing. | Export: write cell values to a text or CSV file while preserving the original workbook.
// AI Prompts: Show a C# Aspose.Cells snippet that iterates the used range and writes each cell’s address and value to a text file. | Provide an example that filters the enumeration to numeric cells only and logs their addresses. | Explain how to modify the loop to output row and column indices instead of the cell’s Name.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, fill sample data, obtain an enumerator for the used range via Cells.GetEnumerator(), iterate each Cell, output its address (Name) and value to the console, and finally save the file.
    public class IterateUsedRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate some sample data (optional, demonstrates the iteration)
                cells["A1"].PutValue("Header");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue(123);
                cells["B2"].PutValue(DateTime.Now);
                cells["C3"].PutValue(true);

                // Get the enumerator for all cells that contain data in the used range
                IEnumerator enumerator = cells.GetEnumerator();

                // Iterate through each cell and log its address (Name) and value
                while (enumerator.MoveNext())
                {
                    Cell cell = (Cell)enumerator.Current;
                    // Log cell address and its value (null check for safety)
                    Console.WriteLine($"{cell.Name}: {cell.Value}");
                }

                // Determine output file path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "IterateUsedRangeDemo.xlsx");

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            IterateUsedRangeDemo.Run();
        }
    }
}
