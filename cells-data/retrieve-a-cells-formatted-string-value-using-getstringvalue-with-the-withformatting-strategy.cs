// Title: Read a formatted cell value (StringValue) from an Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens a .xlsx workbook with Aspose.Cells, accesses a specific cell, and returns its displayed text using the StringValue property that respects the cell's number format. | Create a robust C# example that checks for the workbook file, loads it with Aspose.Cells, retrieves cell B2, obtains the formatted string via StringValue, and handles possible exceptions.
// Common Searches: Aspose.Cells C# get cell displayed text with formatting | How to use StringValue to read formatted value from Excel cell in .NET | Retrieve formatted number from Excel using Aspose.Cells StringValue property | C# Aspose.Cells read cell B2 formatted string from workbook
// Tags: Aspose.Cells retrieve formatted cell string | C# load workbook and read cell display text | Excel number format handling with Aspose.Cells | Get cell B2 formatted value .NET

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that "input.xlsx" exists, loads it into an Aspose.Cells Workbook, accesses the first worksheet, retrieves cell B2, obtains its formatted string via the StringValue property (which respects the cell's formatting), prints the result, and gracefully handles any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve cell B2
            Cell cell = sheet.Cells["B2"];

            // Get the cell's formatted string value (StringValue respects cell formatting)
            string formattedValue = cell.StringValue;

            Console.WriteLine($"Formatted value of B2: {formattedValue}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
