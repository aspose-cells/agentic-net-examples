// Title: Retrieve a cell's raw numeric string using Aspose.Cells GetStringValue with the WithoutFormatting option in C#
// Common Searches: C# Aspose.Cells GetStringValue without formatting to get raw numeric string | How to extract unformatted numeric value from an Excel cell using Aspose.Cells | Aspose.Cells read cell A1 as raw string ignoring number format in .NET | GetStringValue WithoutFormatting example for Excel numeric cells in C# | Retrieve original numeric text from Excel using Aspose.Cells API
// Tags: Aspose.Cells GetStringValue WithoutFormatting | extract raw numeric string from Excel cell C# | unformatted cell value retrieval Aspose.Cells | read numeric cell as string .NET | Excel cell raw text extraction using Aspose API

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel workbook, accesses the first worksheet and cell A1, and extracts the cell's raw numeric string by converting its Value to a string, providing the unformatted numeric text.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the target cell (e.g., A1)
            Cell cell = worksheet.Cells["A1"];

            // Obtain the cell's raw numeric string without any formatting
            string rawNumericString = cell.Value?.ToString() ?? string.Empty;

            // Display the result
            Console.WriteLine($"Raw numeric string: {rawNumericString}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the worksheet: {ex.Message}");
        }
    }
}
