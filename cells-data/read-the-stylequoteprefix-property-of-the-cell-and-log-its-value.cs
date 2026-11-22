// Title: Read the QuotePrefix flag of a cell’s style using Aspose.Cells for .NET in C#
// Common Searches: how to read the QuotePrefix flag from a cell style with Aspose.Cells in C# | Aspose.Cells C# example to check if a cell has a leading quote character | retrieve QuotePrefix property of an Excel cell using Aspose.Cells API | C# code to get cell style QuotePrefix value from a workbook | determine whether Excel cell A1 has QuotePrefix using Aspose.Cells
// Tags: Aspose.Cells read cell QuotePrefix | GetStyle QuotePrefix property C# | Excel cell leading quote detection Aspose.Cells | C# retrieve cell style flag Aspose.Cells | Aspose.Cells style flag extraction example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program loads an Excel workbook, accesses the first worksheet, retrieves the style of cell A1, reads its QuotePrefix boolean property, and writes the value to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file
            string filePath = "Sample.xlsx";

            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(filePath);

                // Get the first worksheet in the workbook
                Worksheet worksheet = workbook.Worksheets[0];

                // Access cell A1
                Cell cell = worksheet.Cells["A1"];

                // Retrieve the cell's style
                Style cellStyle = cell.GetStyle();

                // Read the QuotePrefix property
                bool quotePrefix = cellStyle.QuotePrefix;

                // Output the result
                Console.WriteLine($"QuotePrefix of cell A1: {quotePrefix}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
