// Title: Open an XLSX workbook from a file path using Aspose.Cells Workbook(string) constructor (C#)
// Description: Shows how to load an existing XLSX file by passing its local path to the Aspose.Cells Workbook(string) constructor and then read the name of the first worksheet.
// Keywords: Aspose.Cells | C# | Workbook constructor | load XLSX | open Excel file | read worksheet name | local file path | Aspose.Cells .NET
// Common Searches: Aspose.Cells open XLSX from disk C# | Workbook(string) constructor example Aspose.Cells | How to read first worksheet name with Aspose.Cells | Load Excel workbook using Aspose.Cells C#
// Developer Intent: Load an existing Excel workbook from a local path and access its worksheets with Aspose.Cells.
// Use Cases: Display the name of the first worksheet after opening the file. | Modify cell values, add rows/columns, or insert charts once the workbook is loaded. | Convert the opened workbook to PDF, CSV, or another format. | Retrieve workbook metadata such as total sheet count or sheet visibility.
// AI Prompts: Provide C# code that opens an XLSX file with Aspose.Cells and prints all worksheet names. | Show how to catch and handle FileNotFoundException when using the Workbook(string) constructor in Aspose.Cells. | Explain the steps to open a workbook from a file path and immediately save it as a PDF using Aspose.Cells C#.

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Shows how to load an existing XLSX file by passing its local path to the Aspose.Cells Workbook(string) constructor and then read the name of the first worksheet.
    class Program
    {
        static void Main()
        {
            // Path to the existing XLSX file
            string filePath = "example.xlsx";

            // Open the workbook from the specified file path using the string constructor
            Workbook workbook = new Workbook(filePath);

            // Example operation: display the name of the first worksheet
            Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);
        }
    }
}
