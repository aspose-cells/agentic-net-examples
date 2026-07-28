// Title: C# – Retrieve and Loop Through the Named Range "SalesData" with Aspose.Cells
// Description: Load an Excel workbook, obtain the named range "SalesData" via Workbook.Worksheets.GetRangeByName, enumerate each cell to display its address and value, and optionally save the modified file. Includes file‑existence check and exception handling.
// Keywords: Aspose.Cells | C# | GetRangeByName | named range | enumerate cells | Excel automation | read cell values | Workbook.Save | range iteration | Excel .NET library
// Common Searches: Aspose.Cells GetRangeByName example C# | how to iterate cells in a named range Aspose.Cells | read values from named range SalesData C# | C# enumerate cells in Excel named range | Aspose.Cells retrieve named range and print values
// Developer Intent: The developer needs to access the "SalesData" named range in an existing workbook and loop through its cells to read each cell’s address and content.
// Use Cases: Extract sales figures from the "SalesData" range to feed a reporting engine. | Validate data consistency inside a named range before publishing the workbook. | Apply conditional formatting or calculations to cells within a specific named range.
// AI Prompts: Generate C# code that updates values in the "SalesData" named range using Aspose.Cells and saves the workbook. | Explain best practices for handling a missing named range and suggest alternative ways to locate a range by address. | Rewrite the enumerator loop as a foreach statement to iterate through cells in a named range.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// Load an Excel workbook, obtain the named range "SalesData" via Workbook.Worksheets.GetRangeByName, enumerate each cell to display its address and value, and optionally save the modified file. Includes file‑existence check and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "SalesData"
            // Use fully qualified name to avoid ambiguity with System.Range
            Aspose.Cells.Range salesRange = workbook.Worksheets.GetRangeByName("SalesData");

            if (salesRange != null)
            {
                // Iterate through each cell in the range using its enumerator
                IEnumerator enumerator = salesRange.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    Cell cell = (Cell)enumerator.Current;
                    Console.WriteLine($"{cell.Name}: {cell.Value}");
                }
            }
            else
            {
                Console.WriteLine("Named range 'SalesData' not found.");
            }

            // Save the workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
