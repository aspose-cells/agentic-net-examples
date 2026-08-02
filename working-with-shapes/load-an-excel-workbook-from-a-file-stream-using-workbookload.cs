// Title: Load an Excel Workbook from a Stream with Aspose.Cells (C#)
// Description: Shows how to create a workbook in memory, write data, save it to a MemoryStream, reset the position, load it using the Workbook(Stream) constructor, verify cell content, and finally write the workbook to a file—all with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | load workbook from stream | MemoryStream Excel | Workbook(Stream) constructor | read Excel from stream .NET | in‑memory Excel processing | Aspose.Cells file stream example | load Excel without disk | Aspose.Cells API
// Common Searches: Aspose.Cells load workbook from MemoryStream C# | How to read an Excel file from a stream using Aspose.Cells | Workbook(Stream) constructor usage example | C# load Excel from stream without saving to disk | Aspose.Cells read uploaded Excel stream
// Developer Intent: Load an Excel workbook directly from a stream into a Workbook object, avoiding temporary files.
// Use Cases: Process an Excel file received via an HTTP upload (IFormFile) entirely in memory. | Read and modify an Excel document streamed from a cloud storage service before saving or converting it. | Convert a streamed Excel workbook to PDF or another format without creating intermediate files on disk.
// AI Prompts: Provide C# code that loads an Excel workbook from a Stream with Aspose.Cells, changes cell B2, and saves to a new file. | Explain why resetting the stream position is required before calling the Workbook(Stream) constructor. | Show how to load an uploaded Excel stream, rename the first worksheet, and export the workbook to PDF using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to create a workbook in memory, write data, save it to a MemoryStream, reset the position, load it using the Workbook(Stream) constructor, verify cell content, and finally write the workbook to a file—all with Aspose.Cells for .NET.
class LoadWorkbookFromStream
{
    static void Main()
    {
        // ------------------------------------------------------------
        // Create a sample workbook in memory and write it to a stream.
        // ------------------------------------------------------------
        MemoryStream memoryStream = new MemoryStream();
        Workbook sampleWorkbook = new Workbook();                     // create a new workbook
        sampleWorkbook.Worksheets[0].Cells["A1"].PutValue("Hello from stream");
        sampleWorkbook.Save(memoryStream, SaveFormat.Xlsx);          // save to the memory stream
        memoryStream.Position = 0;                                   // reset stream position for reading

        // ------------------------------------------------------------
        // Load the workbook from the stream using the Workbook(Stream) ctor.
        // ------------------------------------------------------------
        Workbook loadedWorkbook = new Workbook(memoryStream);        // load from stream

        // ------------------------------------------------------------
        // Demonstrate that the data was loaded correctly.
        // ------------------------------------------------------------
        string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine($"Loaded cell A1 value: {cellValue}");

        // ------------------------------------------------------------
        // Save the loaded workbook to a physical file.
        // ------------------------------------------------------------
        loadedWorkbook.Save("LoadedFromStream.xlsx", SaveFormat.Xlsx);
    }
}
