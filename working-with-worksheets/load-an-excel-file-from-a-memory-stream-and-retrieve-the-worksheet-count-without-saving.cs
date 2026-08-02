// Title: C# – Load Excel Workbook from MemoryStream & Get Worksheet Count using Aspose.Cells
// Description: Demonstrates how to create a workbook, save it to a MemoryStream, load it directly from the stream with the Workbook(Stream) constructor, and read the Worksheets.Count property—all without writing any files to disk.
// Keywords: Aspose.Cells | C# | .NET | MemoryStream | load workbook from stream | worksheet count | in‑memory Excel | stream constructor | XLSX without file system
// Common Searches: Aspose.Cells load workbook from MemoryStream C# | Get number of worksheets from Excel stream Aspose | Read Excel file without saving to disk .NET | Count sheets in in‑memory workbook Aspose.Cells | Workbook(Stream) constructor example
// Developer Intent: Load an Excel file from a MemoryStream and determine how many worksheets it contains without persisting the file.
// Use Cases: Validate the structure of an uploaded Excel file in a web API by counting sheets directly from the request stream. | Process a dynamically generated Excel report stored in a byte array, retrieve the sheet count, and continue in‑memory manipulation. | Integrate Aspose.Cells into server‑side services where disk I/O is restricted, such as sandboxed environments or cloud functions.
// AI Prompts: Write C# code that uses Aspose.Cells to read an Excel workbook from a MemoryStream and return the total worksheet count. | Show how to reset a MemoryStream, load a Workbook via the stream constructor, rename the first sheet, and output the worksheet count. | Explain best practices for handling large Excel files entirely in memory with Aspose.Cells, including loading from a stream and efficiently obtaining the sheet count.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates how to create a workbook, save it to a MemoryStream, load it directly from the stream with the Workbook(Stream) constructor, and read the Worksheets.Count property—all without writing any files to disk.
    class LoadWorkbookFromMemoryStream
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // Step 1: Create a sample workbook and save it to a MemoryStream
            // ------------------------------------------------------------
            Workbook originalWorkbook = new Workbook();               // Create a new workbook (default has one worksheet)
            originalWorkbook.Worksheets[0].Name = "SampleSheet";    // Optional: rename the default sheet
            originalWorkbook.Worksheets[0].Cells["A1"].PutValue("Hello, Aspose!"); // Add some data

            // Save the workbook into a memory stream in XLSX format
            using (MemoryStream memoryStream = new MemoryStream())
            {
                originalWorkbook.Save(memoryStream, SaveFormat.Xlsx);
                memoryStream.Position = 0; // Reset stream position for reading

                // ------------------------------------------------------------
                // Step 2: Load a workbook from the same memory stream
                // ------------------------------------------------------------
                Workbook loadedWorkbook = new Workbook(memoryStream); // Load using the Stream constructor

                // ------------------------------------------------------------
                // Step 3: Retrieve and display the worksheet count
                // ------------------------------------------------------------
                int worksheetCount = loadedWorkbook.Worksheets.Count;
                Console.WriteLine($"Number of worksheets in the loaded workbook: {worksheetCount}");
            }

            // No saving to disk is performed, as per the requirement.
        }
    }
}
