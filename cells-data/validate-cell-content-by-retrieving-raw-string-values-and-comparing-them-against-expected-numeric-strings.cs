// Title: Validate raw StringValue of specific Excel cells against expected numeric strings using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, reads the StringValue of specified cell coordinates, and checks each value against a predefined numeric string. | Create a loop that iterates over a list of (row, column, expected) tuples, retrieves each cell's raw text, and prints whether it matches the expected string using ordinal comparison.
// Common Searches: Aspose.Cells C# get unformatted cell text StringValue example | how to compare Excel cell raw value to a numeric string in .NET | validate multiple cell values programmatically with Aspose.Cells
// Tags: retrieve raw StringValue Aspose.Cells | compare cell raw text to expected numeric string | validate multiple cells Aspose.Cells .NET | ordinal string comparison Excel cell value | worksheet cell verification Aspose.Cells

using Aspose.Cells;
using System;

// The program loads an Excel workbook using Aspose.Cells, iterates over a predefined array of cell coordinates, reads each cell's raw StringValue (ignoring any formatting), compares it to an expected numeric string with ordinal comparison, and outputs the validation result to the console.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the cells to validate: row index, column index, and the expected numeric string
        var cellsToValidate = new (int Row, int Column, string Expected)[]
        {
            (0, 0, "123"), // Cell A1
            (1, 1, "456")  // Cell B2
        };

        // Iterate through each cell definition
        foreach (var (row, column, expected) in cellsToValidate)
        {
            // Retrieve the cell object
            Cell cell = worksheet.Cells[row, column];

            // Get the raw string value stored in the cell (no formatting applied)
            string rawValue = cell.StringValue;

            // Compare the raw string with the expected numeric string
            bool isValid = string.Equals(rawValue, expected, StringComparison.Ordinal);

            // Output the validation result
            Console.WriteLine($"Cell {cell.Name}: Raw='{rawValue}' Expected='{expected}' Valid={isValid}");
        }
    }
}
