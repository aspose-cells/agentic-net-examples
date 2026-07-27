// Title: Open an XLSX workbook from a local path with Aspose.Cells Workbook(string) in C#
// Description: Demonstrates how to instantiate a Workbook by passing a full file‑system path to the Workbook(string) constructor, retrieve the first worksheet, and print its name and the value of cell A1 to the console, confirming that the XLSX file is loaded correctly.
// Keywords: Aspose.Cells | C# Workbook(string) | open XLSX file | load workbook from path | .NET read Excel | first worksheet name | cell A1 value
// Common Searches: Aspose.Cells open existing XLSX C# | Workbook(string) constructor example | C# read Excel file path Aspose.Cells | how to get worksheet name Aspose.Cells | retrieve cell A1 value after loading workbook
// Developer Intent: Load an existing XLSX file from a specified local path and access its first worksheet.
// Use Cases: Load a template workbook, populate data programmatically, and save the result. | Read configuration values stored in the first sheet before executing business logic. | Validate that a required worksheet exists and that cell A1 contains a header before processing rows.
// AI Prompts: Generate C# code that opens an XLSX file using Aspose.Cells Workbook(string) and prints the first worksheet name and cell A1 value. | Add comprehensive error handling for missing files, empty workbooks, or absent worksheets when using the Workbook(string) constructor. | Show how to open a workbook from a relative path and iterate over all cells in the first row with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsOpenExample
{
    // Demonstrates how to instantiate a Workbook by passing a full file‑system path to the Workbook(string) constructor, retrieve the first worksheet, and print its name and the value of cell A1 to the console, confirming that the XLSX file is loaded correctly.
    class Program
    {
        static void Main()
        {
            // Path to the existing XLSX file on the local file system
            string filePath = @"C:\Data\SampleWorkbook.xlsx";

            // Open the workbook using the constructor that accepts a file path
            // This follows the provided rule: Workbook(string)
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet to demonstrate that the file was loaded
            Worksheet sheet = workbook.Worksheets[0];

            // Output the name of the first worksheet and the value of cell A1 (if any)
            Console.WriteLine($"Worksheet Name: {sheet.Name}");
            Console.WriteLine($"Cell A1 Value: {sheet.Cells["A1"].StringValue}");
        }
    }
}
