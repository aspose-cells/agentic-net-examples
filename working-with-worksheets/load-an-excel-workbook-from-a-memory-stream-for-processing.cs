// Title: Load an Excel workbook from a MemoryStream using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, write data to cells, save it to a MemoryStream in XLSX format, reset the stream, and instantiate a new Workbook directly from that MemoryStream to read cell values without touching the file system.
// Keywords: Aspose.Cells | C# | .NET | MemoryStream | load workbook from stream | Workbook(Stream) constructor | save workbook to stream | in‑memory Excel processing | read Excel cells without file | avoid disk I/O
// Common Searches: Aspose.Cells load workbook from MemoryStream C# | Workbook(Stream) example Aspose.Cells | read Excel file from stream .NET | save Excel to MemoryStream Aspose | process Excel in memory without file
// Developer Intent: Load an Excel workbook directly from a MemoryStream to read or manipulate its contents without creating a physical file.
// Use Cases: Handle an uploaded Excel file in an ASP.NET controller by streaming it into memory for validation and data extraction. | Receive Excel data over a network socket, store it in a MemoryStream, and instantiate a Workbook for quick parsing. | Convert an in‑memory workbook to another format (e.g., PDF) after loading it from a stream, eliminating temporary files.
// AI Prompts: Generate C# code that loads an Excel file from a MemoryStream with Aspose.Cells, updates a cell, and saves the result to a PDF stream. | Explain how to process large Excel files with Aspose.Cells by streaming them from MemoryStream to minimize memory usage. | Provide an example that loads multiple workbooks from separate MemoryStreams in a loop and aggregates data from each workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryLoadDemo
{
    // Demonstrates how to create a workbook, write data to cells, save it to a MemoryStream in XLSX format, reset the stream, and instantiate a new Workbook directly from that MemoryStream to read cell values without touching the file system.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a workbook and add some sample data.
            // -------------------------------------------------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Cells["A1"].PutValue("Sample");
            sourceSheet.Cells["B1"].PutValue(123);

            // -------------------------------------------------
            // 2. Save the workbook into a MemoryStream.
            //    Using the Save method that writes to a stream.
            // -------------------------------------------------
            using (MemoryStream memoryStream = new MemoryStream())
            {
                // Save as XLSX format into the stream.
                sourceWorkbook.Save(memoryStream, SaveFormat.Xlsx);

                // Reset the stream position to the beginning before reading.
                memoryStream.Position = 0;

                // -------------------------------------------------
                // 3. Load a new workbook from the same MemoryStream.
                //    This uses the Workbook(Stream) constructor rule.
                // -------------------------------------------------
                Workbook loadedWorkbook = new Workbook(memoryStream);

                // -------------------------------------------------
                // 4. Process the loaded workbook (example: read a cell).
                // -------------------------------------------------
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                Console.WriteLine("Cell A1 value: " + loadedSheet.Cells["A1"].StringValue);
                Console.WriteLine("Cell B1 value: " + loadedSheet.Cells["B1"].IntValue);
            }
        }
    }
}
