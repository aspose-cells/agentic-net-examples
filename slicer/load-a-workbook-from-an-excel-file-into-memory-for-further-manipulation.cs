// Title: C# – Load an Excel workbook with Aspose.Cells and read cell A1
// Description: Demonstrates how to verify an Excel file's presence, create a Workbook object from the file path using Aspose.Cells for .NET, access the first worksheet, and output the value of cell A1 while handling possible exceptions.
// Keywords: Aspose.Cells load workbook C# | read Excel cell Aspose.Cells | verify file exists before opening Excel | Workbook object Aspose.Cells .NET | exception handling Aspose.Cells
// Common Searches: load Excel file into Aspose.Cells workbook C# | read cell A1 after opening workbook with Aspose.Cells | check if Excel file exists before creating Workbook in .NET | Aspose.Cells example for opening and reading Excel
// Developer Intent: Open an existing Excel file, create a Workbook instance, and retrieve a cell value.
// Use Cases: Load a workbook from a known path and read specific cell data. | Prevent FileNotFoundException by confirming the file exists before loading. | Initialize the first worksheet for further data processing after opening the workbook.
// AI Prompts: Generate C# code that uses Aspose.Cells to open an Excel file, verify its existence, and print the value of cell B2. | Create an Aspose.Cells example that loads a workbook, catches errors, and iterates over all cells in the first row.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to verify an Excel file's presence, create a Workbook object from the file path using Aspose.Cells for .NET, access the first worksheet, and output the value of cell A1 while handling possible exceptions.
    public class LoadWorkbookExample
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the Excel file to be loaded
            string filePath = "example.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook from the file into memory
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet for further manipulation
                Worksheet worksheet = workbook.Worksheets[0];

                // Example operation: read and display the value of cell A1
                Console.WriteLine("Cell A1 value: " + worksheet.Cells["A1"].StringValue);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
