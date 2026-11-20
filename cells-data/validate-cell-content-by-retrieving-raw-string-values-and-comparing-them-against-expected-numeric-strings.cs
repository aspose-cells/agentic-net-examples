// Title: Validate raw string values of specific Excel cells against expected numeric strings with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# retrieve cell raw string without formatting | How to compare an Excel cell's string value to a numeric string in .NET | Validate contents of cells A1 and B2 in an XLSX file using Aspose.Cells | Check if a worksheet cell equals a predefined string with Aspose.Cells C# | Extract unformatted cell text and verify it against expected data in C#
// Tags: retrieve raw cell string Aspose.Cells | compare cell value to expected string C# | validate specific Excel cells Aspose.Cells | unformatted cell text extraction .NET | Excel cell content verification Aspose.Cells | load workbook and check cell values C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX workbook with Aspose.Cells, reads the raw (unformatted) string values of designated cells (e.g., A1 and B2), compares each to a predefined numeric string using ordinal comparison, outputs validation results, and saves the workbook while handling missing files and password‑protected workbooks.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cells to validate and their expected numeric strings
            var validations = new (int Row, int Column, string Expected)[]
            {
                (0, 0, "123"),   // Cell A1
                (1, 1, "456")    // Cell B2
            };

            // Iterate through each validation rule
            foreach (var v in validations)
            {
                // Retrieve the cell at the specified row and column
                Cell cell = sheet.Cells[v.Row, v.Column];

                // Get the raw string value stored in the cell (no formatting applied)
                string rawValue = cell.StringValue;

                // Compare the raw string with the expected numeric string
                bool isValid = string.Equals(rawValue, v.Expected, StringComparison.Ordinal);

                // Output the result
                Console.WriteLine($"Cell {cell.Name}: Raw='{rawValue}' Expected='{v.Expected}' Valid={isValid}");
            }

            // Save the workbook (optional, if any changes were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle password-protected files or other errors
            if (ex.Message.Contains("password", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Error: The workbook is password protected. Provide a valid password to open it.");
            }
            else
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
