// Title: C# – Load an Excel Workbook from Disk and Access the Initial Worksheet with Aspose.Cells
// Description: Shows how to create a Workbook from a file path, retrieve the worksheet at index 0, and print its Name property. Perfect for quickly opening existing .xlsx files in .NET projects.
// Keywords: Aspose.Cells load workbook C# | open Excel file Aspose.Cells | initial worksheet index 0 | Workbook(string) constructor | read worksheet name C# | Aspose.Cells .NET example | load .xlsx from file
// Common Searches: Aspose.Cells open existing .xlsx C# | Get first sheet name Aspose.Cells | C# load Excel workbook from path | How to read worksheet name using Aspose.Cells | Aspose.Cells workbook constructor example
// Developer Intent: Open a saved Excel file and obtain a reference to its initial worksheet.
// Use Cases: Validate a template by confirming the initial sheet's title matches expectations. | Process user‑uploaded spreadsheets where the data begins on the first tab. | Start reporting logic that always uses the workbook's primary worksheet.
// AI Prompts: Write C# code that uses Aspose.Cells to open a given .xlsx file and return the name of the initial worksheet. | Show how to add error handling for missing or corrupted Excel files when creating a Workbook with Aspose.Cells. | Provide a snippet that loads a workbook and then loops through all worksheets to print their names.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    // Shows how to create a Workbook from a file path, retrieve the worksheet at index 0, and print its Name property. Perfect for quickly opening existing .xlsx files in .NET projects.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string filePath = "input.xlsx";

            // Load the workbook from the specified file using the Workbook(string) constructor
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet in the workbook (index 0)
            Worksheet firstWorksheet = workbook.Worksheets[0];

            // Example usage: display the name of the first worksheet
            Console.WriteLine($"First worksheet name: {firstWorksheet.Name}");
        }
    }
}
