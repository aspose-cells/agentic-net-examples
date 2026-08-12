// Title: C# – Load an Excel Workbook (XLS or XLSX) with Aspose.Cells using the Constructor or Load Method
// Description: Demonstrates how to instantiate an Aspose.Cells Workbook from a local .xls or .xlsx file, access the first worksheet, read the value of cell A1, output it to the console, and properly dispose the workbook in a .NET application.
// Keywords: Aspose.Cells load workbook C# | open Excel file .NET | Workbook constructor Aspose.Cells | Workbook.Load method example | read cell A1 Aspose.Cells | C# Excel file handling | dispose Aspose.Cells workbook | XLSX to console Aspose.Cells
// Common Searches: How to open an .xlsx file with Aspose.Cells in C# | Aspose.Cells read cell value after loading workbook | Workbook.Load vs constructor Aspose.Cells | C# example for loading Excel workbook using Aspose.Cells | Dispose Aspose.Cells workbook after use
// Developer Intent: Load an Excel file into an Aspose.Cells Workbook and retrieve cell data.
// Use Cases: Quickly open a spreadsheet to extract a single cell value (e.g., A1). | Load a workbook for further processing such as iterating worksheets or rows. | Release unmanaged resources by disposing the Workbook after operations.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells using Workbook.Load and prints the value of cell B2. | Show how to catch and log exceptions when a corrupted .xls file is loaded with Aspose.Cells. | Create a reusable method that opens any Excel file with Aspose.Cells and returns a dictionary of cell addresses and values.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    // Demonstrates how to instantiate an Aspose.Cells Workbook from a local .xls or .xlsx file, access the first worksheet, read the value of cell A1, output it to the console, and properly dispose the workbook in a .NET application.
    class Program
    {
        static void Main()
        {
            // Path to the Excel file to be loaded (can be .xls or .xlsx)
            string filePath = "sample.xlsx";

            // Load the workbook using the constructor that accepts a file path
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Read and display the value of cell A1
            Console.WriteLine("Value of A1: " + worksheet.Cells["A1"].StringValue);

            // Optional: clean up resources
            workbook.Dispose();
        }
    }
}
