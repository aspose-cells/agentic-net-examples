// Title: Read‑only enumeration of worksheet cells to collect all numeric values using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx workbook with Aspose.Cells, uses a read‑only Cells enumerator to walk through each cell in the first worksheet, and appends every numeric cell's value to a List<double>. | Provide a C# snippet that checks for the presence of an input Excel file, loads it via Aspose.Cells, iterates over its cells without altering them, and gathers all numbers into a collection while handling possible errors.
// Common Searches: Aspose.Cells example for read‑only cell iteration to retrieve numbers in C# | How to gather all numeric data from an Excel sheet using Aspose.Cells .NET | C# enumerate worksheet cells without modification Aspose.Cells | Extract double values from .xlsx file with Aspose.Cells library
// Tags: Aspose.Cells read‑only cell enumeration | collect numeric cells C# | iterate worksheet cells without changes | extract numeric data from Excel Aspose.Cells | check input file existence C#

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The program verifies that input.xlsx exists, loads it with Aspose.Cells, iterates over every cell in the first worksheet using a read‑only Cells enumerator, identifies numeric cells, converts their values to double, and stores them in a List<double>.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // List to collect all numeric values
            List<double> numericValues = new List<double>();

            // Iterate through all cells using foreach (generic enumerator)
            foreach (Cell cell in sheet.Cells)
            {
                // Check if the cell contains a numeric value
                if (cell.Type == CellValueType.IsNumeric && cell.Value != null)
                {
                    // Convert the value to double and store it
                    numericValues.Add(Convert.ToDouble(cell.Value));
                }
            }

            // Example output: number of numeric values found
            Console.WriteLine($"Numeric values collected: {numericValues.Count}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
