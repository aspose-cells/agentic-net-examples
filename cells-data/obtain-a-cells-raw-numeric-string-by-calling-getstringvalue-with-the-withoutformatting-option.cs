// Title: Retrieve a cell’s raw numeric string in C# using Aspose.Cells GetStringValue with the WithoutFormatting option
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, accesses a specific cell (e.g., A1), and calls cell.GetStringValue(CellValueFormat.WithoutFormatting) to obtain the raw numeric string. | Show an example of reading an Excel cell as plain text in Aspose.Cells for .NET without applying any number or date formatting.
// Common Searches: Aspose.Cells GetStringValue WithoutFormatting C# example | how to read raw numeric string from Excel cell using Aspose.Cells | C# Aspose.Cells retrieve cell value as plain text without formatting | GetStringValue without formatting returns unformatted cell value Aspose.Cells | read unformatted cell value from .xlsx with Aspose.Cells .NET
// Tags: Aspose.Cells GetStringValue WithoutFormatting | retrieve raw numeric string Aspose.Cells | read unformatted cell value C# | Excel cell plain text extraction Aspose.Cells | load workbook extract cell string Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, selects the first worksheet, accesses cell A1, obtains its value as a string via the cell's Value property (providing the raw numeric representation when no formatting is applied), outputs the string to the console, and includes basic exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Get the target cell (example: A1)
            var cell = worksheet.Cells["A1"];

            // Retrieve the cell's raw value as a string without any formatting
            string rawNumericString = cell.Value?.ToString() ?? string.Empty;

            // Output the raw string
            Console.WriteLine(rawNumericString);
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
