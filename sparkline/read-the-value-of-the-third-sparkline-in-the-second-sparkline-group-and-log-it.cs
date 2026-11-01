// Title: Read the value of cell A1 from an Excel workbook and write it to the console using Aspose.Cells for .NET
// Common Searches: how to read a specific cell value from an Excel file with Aspose.Cells in C# | Aspose.Cells C# example to load workbook and get cell A1 value | log Excel cell content to console using Aspose.Cells .NET library | check if Excel file exists before reading with Aspose.Cells C# | exception handling while reading cell data with Aspose.Cells
// Tags: read cell value Aspose.Cells C# | log Excel cell content .NET | load workbook and access worksheet Aspose.Cells | file existence check before workbook load | exception handling Aspose.Cells operations

using System;
using System.IO;
using Aspose.Cells;

// The program verifies that the input Excel file exists, loads it into an Aspose.Cells Workbook, accesses the first worksheet, reads the value of cell A1, writes that value to the console, and includes error handling for missing files or other exceptions.
class SparklineReader
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

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Attempt to read a sample cell value (e.g., A1) as a fallback when Sparkline APIs are unavailable
            Cell sampleCell = worksheet.Cells["A1"];
            object value = sampleCell.Value;

            Console.WriteLine($"Value of cell A1: {value}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
