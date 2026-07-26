// Title: C# – Load an Excel file by path and get the first worksheet with Aspose.Cells
// Description: Shows how to create a Workbook from a file path, access the first worksheet (index 0) and print its name using Aspose.Cells for .NET.
// Keywords: Aspose.Cells load workbook C# | open Excel file Aspose.Cells | first worksheet index Aspose.Cells | read worksheet name C# | Aspose.Cells .NET example | load workbook from file path
// Common Searches: Aspose.Cells open existing workbook C# | how to read first sheet name Aspose.Cells | C# Aspose.Cells get worksheet by index | load Excel file from disk Aspose.Cells
// Developer Intent: Open an existing Excel workbook and obtain a reference to its first worksheet.
// Use Cases: Display the name of the first sheet after a user uploads an Excel file. | Iterate rows of the first worksheet to import data into a database. | Copy the first worksheet into a new workbook for a custom report. | Validate that the expected sheet exists before processing.
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells and returns the name of the first worksheet. | Show how to handle FileNotFoundException and other I/O errors when loading a workbook with Aspose.Cells. | Demonstrate loading a workbook from a MemoryStream instead of a file path using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to create a Workbook from a file path, access the first worksheet (index 0) and print its name using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Specify the path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook from the given file path using the string constructor
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet in the workbook (zero‑based index)
        Worksheet firstWorksheet = workbook.Worksheets[0];

        // Example usage: output the name of the first worksheet
        Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
    }
}
