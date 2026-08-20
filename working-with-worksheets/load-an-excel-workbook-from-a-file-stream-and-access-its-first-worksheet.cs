// Title: Load Excel Workbook from FileStream and Access First Worksheet – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to open an Excel file with `File.OpenRead`, instantiate a `Workbook` using the `Workbook(Stream)` constructor, and retrieve the first worksheet via the `Worksheets` collection. The sample prints the worksheet name to the console.
// Keywords: Aspose.Cells | C# load workbook from stream | Workbook(Stream) constructor | FileStream Excel .NET | first worksheet access | read Excel file Aspose.Cells | open Excel from MemoryStream | Aspose.Cells example
// Common Searches: Aspose.Cells open Excel from FileStream C# | Workbook(Stream) constructor usage | Get first worksheet after loading workbook Aspose.Cells | Read Excel file with Aspose.Cells .NET | How to load Excel workbook from stream in C#
// Developer Intent: Open an Excel workbook via a FileStream and retrieve its first worksheet using Aspose.Cells in C#.
// Use Cases: Display the name of the first worksheet after a user uploads an Excel file to a web service. | Iterate rows of the first sheet for data import when processing large Excel files from a stream. | Expose worksheet metadata (name, index, visibility) through an API without saving the file to disk.
// AI Prompts: Generate C# code that loads an Excel workbook from a MemoryStream and prints the first worksheet name using Aspose.Cells. | Show how to add robust error handling when opening an Excel file from a FileStream with Aspose.Cells. | Provide an example that reads cell A1 of the first worksheet after loading a workbook from a stream in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Demonstrates how to open an Excel file with `File.OpenRead`, instantiate a `Workbook` using the `Workbook(Stream)` constructor, and retrieve the first worksheet via the `Worksheets` collection. The sample prints the worksheet name to the console.
    class LoadWorkbookFromStream
    {
        static void Main()
        {
            // Path to the Excel file to be loaded
            string filePath = "input.xlsx";

            // Open a file stream for reading the Excel file
            using (FileStream stream = File.OpenRead(filePath))
            {
                // Load the workbook from the opened stream using the Workbook(Stream) constructor
                Workbook workbook = new Workbook(stream);

                // Access the first worksheet in the workbook
                Worksheet firstWorksheet = workbook.Worksheets[0];

                // Example usage: output the name of the first worksheet
                Console.WriteLine($"First worksheet name: {firstWorksheet.Name}");
            }
        }
    }
}
