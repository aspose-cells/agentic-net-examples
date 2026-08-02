// Title: Load an Excel workbook from a FileStream and get the first worksheet – Aspose.Cells for .NET (C#)
// Description: This example opens an Excel file as a read‑only FileStream, creates a Workbook using the Workbook(Stream) constructor, accesses the first worksheet (index 0), reads cell A1 and writes the value to the console.
// Keywords: Aspose.Cells load workbook stream | C# open Excel FileStream | first worksheet Aspose.Cells | read cell A1 Aspose.Cells | Workbook(Stream) constructor | stream‑based Excel loading | Aspose.Cells read‑only stream | C# Excel file stream example
// Common Searches: Aspose.Cells open Excel file from FileStream C# | How to get the first worksheet after loading a workbook from a stream | Read cell values from a workbook loaded via stream using Aspose.Cells | C# example for Workbook(Stream) constructor Aspose.Cells
// Developer Intent: Load an Excel workbook from a stream and work with its first worksheet.
// Use Cases: Read specific cells (e.g., A1) from the first worksheet after stream‑based loading. | Extract tabular data by iterating rows of the first worksheet without saving the file to disk. | Apply formatting, formulas, or data validation to the first worksheet immediately after opening the workbook from a stream. | Integrate stream‑loaded workbooks into web APIs or services where files are received as streams.
// AI Prompts: Generate C# code that loads an Excel workbook from a MemoryStream and returns the first worksheet using Aspose.Cells. | Show how to read multiple cell values from the first worksheet after opening an Excel file as a read‑only FileStream with Aspose.Cells. | Explain best practices for exception handling and resource disposal when loading a workbook from a stream in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // This example opens an Excel file as a read‑only FileStream, creates a Workbook using the Workbook(Stream) constructor, accesses the first worksheet (index 0), reads cell A1 and writes the value to the console.
    class Program
    {
        static void Main()
        {
            // Path to the Excel file to be loaded
            string filePath = "sample.xlsx";

            // Open the file as a read‑only stream
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Load the workbook from the stream using the Workbook(Stream) constructor
                Workbook workbook = new Workbook(fileStream);

                // Access the first worksheet (index 0)
                Worksheet firstWorksheet = workbook.Worksheets[0];

                // Example: read the value of cell A1 and display it
                string cellValue = firstWorksheet.Cells["A1"].StringValue;
                Console.WriteLine($"Value in A1: {cellValue}");
            }
        }
    }
}
