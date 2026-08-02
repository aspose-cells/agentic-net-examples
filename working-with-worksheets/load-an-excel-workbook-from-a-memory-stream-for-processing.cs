// Title: Load an Excel workbook from a MemoryStream using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, write data, save it to a MemoryStream (Excel 97‑2003 format), reset the stream, reload the workbook from the stream, read cell values, display them, and optionally save the processed file as XLSX—all without touching the file system.
// Keywords: Aspose.Cells MemoryStream | load workbook from stream C# | read Excel from memory Aspose | save Aspose.Cells to stream | in‑memory Excel processing .NET | C# Aspose.Cells example | Excel 97-2003 to XLSX conversion
// Common Searches: Aspose.Cells load workbook from MemoryStream | C# read Excel data without saving to disk | How to reset MemoryStream position before loading Aspose.Cells | Convert Excel 97-2003 stream to XLSX using Aspose.Cells | Aspose.Cells example for streaming workbook
// Developer Intent: Load and manipulate an Excel workbook directly from a MemoryStream to avoid file‑system I/O, then read or modify its contents programmatically.
// Use Cases: Process Excel files received from a web API or service without writing temporary files. | Transmit a workbook over a network as a byte stream, reconstruct it on the client side, and apply calculations or formatting. | Perform high‑performance batch operations on large workbooks by streaming them in memory rather than using disk storage.
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, saves it to a MemoryStream, resets the stream, reloads the workbook, and prints each cell value. | Write a method that accepts a MemoryStream containing an Excel file and returns a list of product names and prices using Aspose.Cells. | Show how to stream a large Excel workbook to a MemoryStream with Aspose.Cells, modify some cells, and then save the result as an XLSX file without intermediate disk writes.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryLoadDemo
{
    // Demonstrates how to create a workbook, write data, save it to a MemoryStream (Excel 97‑2003 format), reset the stream, reload the workbook from the stream, read cell values, display them, and optionally save the processed file as XLSX—all without touching the file system.
    public class Program
    {
        public static void Main()
        {
            // 1. Create a workbook and add sample data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sheet = sourceWorkbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue(999.99);
            sheet.Cells["A3"].PutValue("Phone");
            sheet.Cells["B3"].PutValue(699.99);

            // 2. Save the workbook to a memory stream (Excel 97-2003 format)
            MemoryStream memoryStream = sourceWorkbook.SaveToStream();

            // 3. Reset the stream position before reading
            memoryStream.Position = 0;

            // 4. Load a new workbook from the memory stream
            Workbook loadedWorkbook = new Workbook(memoryStream);

            // 5. Process the loaded workbook (example: read and display cell values)
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Console.WriteLine("Loaded Workbook Content:");
            for (int row = 0; row <= 2; row++)
            {
                string product = loadedSheet.Cells[row, 0].StringValue;
                string price = loadedSheet.Cells[row, 1].StringValue;
                Console.WriteLine($"{product}\t{price}");
            }

            // Optional: Save the processed workbook to a file
            loadedWorkbook.Save("ProcessedFromMemory.xlsx", SaveFormat.Xlsx);
        }
    }
}
