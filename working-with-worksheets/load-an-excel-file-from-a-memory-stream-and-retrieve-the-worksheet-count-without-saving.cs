// Title: Load Excel workbook from MemoryStream and get worksheet count with Aspose.Cells for .NET
// Description: Creates a workbook with two sheets, saves it to a MemoryStream (XLSX), resets the stream, loads a new Workbook directly from the stream, and reads Worksheets.Count—all without writing to disk.
// Keywords: Aspose.Cells | .NET | C# | MemoryStream | load workbook from stream | worksheet count | in‑memory Excel | Workbook.SaveToStream | Workbook(Stream) constructor | no file system I/O
// Common Searches: Aspose.Cells load workbook from memory stream | C# count worksheets in Excel file without saving | read Excel from byte array using Aspose.Cells | get worksheet count from stream Aspose | in‑memory Excel processing Aspose.Cells .NET
// Developer Intent: Load an Excel file from a MemoryStream and determine the number of worksheets without persisting the file to disk.
// Use Cases: Validate sheet count of an Excel payload received via a web API by loading the byte array into a MemoryStream and checking Worksheets.Count. | Generate a report in memory, serialize it to a stream, reload it to verify the worksheet structure before sending the file to a client. | Run a background batch that processes many workbooks entirely in memory to count sheets and avoid costly file‑system operations.
// AI Prompts: Provide C# code using Aspose.Cells that reads an Excel file from a MemoryStream and returns the total number of worksheets. | Show an example that creates a workbook, saves it to a MemoryStream, reloads it with the Workbook(Stream) constructor, and prints the worksheet count. | Explain why resetting the MemoryStream position is required before loading it with Aspose.Cells and how to do it correctly.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook with two sheets, saves it to a MemoryStream (XLSX), resets the stream, loads a new Workbook directly from the stream, and reads Worksheets.Count—all without writing to disk.
class LoadFromMemoryStreamDemo
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a sample workbook with two worksheets
        // ------------------------------------------------------------
        Workbook originalWorkbook = new Workbook();               // uses Workbook() constructor
        originalWorkbook.Worksheets[0].Name = "FirstSheet";

        // Add a second worksheet
        int secondSheetIndex = originalWorkbook.Worksheets.Add();
        originalWorkbook.Worksheets[secondSheetIndex].Name = "SecondSheet";

        // ------------------------------------------------------------
        // 2. Save the workbook to a MemoryStream (Excel 2007+ format)
        // ------------------------------------------------------------
        // SaveToStream returns a MemoryStream containing the XLSX data
        MemoryStream memoryStream = originalWorkbook.SaveToStream(); // uses Workbook.SaveToStream()

        // Reset the stream position to the beginning before reading
        memoryStream.Position = 0;

        // ------------------------------------------------------------
        // 3. Load a new workbook from the MemoryStream
        // ------------------------------------------------------------
        Workbook loadedWorkbook = new Workbook(memoryStream);    // uses Workbook(Stream) constructor

        // ------------------------------------------------------------
        // 4. Retrieve and display the worksheet count
        // ------------------------------------------------------------
        int worksheetCount = loadedWorkbook.Worksheets.Count;
        Console.WriteLine($"Worksheet count loaded from memory stream: {worksheetCount}");

        // Clean up
        memoryStream.Dispose();
        originalWorkbook.Dispose();
        loadedWorkbook.Dispose();
    }
}
