// Title: Load an Excel workbook from a MemoryStream, edit cell A1, and save it back to a MemoryStream with Aspose.Cells for .NET
// Common Searches: C# load Excel file from MemoryStream using Aspose.Cells | update a specific cell in an in‑memory workbook with Aspose.Cells | save modified Aspose.Cells workbook to a MemoryStream as XLSX | Aspose.Cells example for reading and writing Excel data without disk | how to reset MemoryStream position after saving workbook in C#
// Tags: load workbook from memory stream Aspose.Cells | modify cell value in worksheet C# | save workbook to memory stream XLSX format | in‑memory Excel processing Aspose.Cells | reset MemoryStream position after save | first worksheet access Aspose.Cells | stream‑based Excel editing .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample demonstrates how to obtain an input MemoryStream containing an XLSX file, load it into an Aspose.Cells Workbook, access the first worksheet, change the value of cell A1, and then save the updated workbook into a new MemoryStream in XLSX format, resetting the stream position for further consumption.
class Program
{
    static void Main()
    {
        // Obtain the input memory stream that contains the Excel file.
        // Replace GetInputStream() with your actual source of the stream.
        MemoryStream inputStream = GetInputStream();

        // Load the workbook from the memory stream.
        Workbook workbook = new Workbook(inputStream);

        // Access the first worksheet.
        Worksheet worksheet = workbook.Worksheets[0];

        // Modify the value of cell A1.
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("Updated Value");

        // Prepare an output memory stream to write the modified workbook.
        MemoryStream outputStream = new MemoryStream();

        // Save the workbook back to the output stream in XLSX format.
        workbook.Save(outputStream, SaveFormat.Xlsx);

        // Reset the position of the output stream to the beginning if it will be read later.
        outputStream.Position = 0;

        // At this point, outputStream contains the updated Excel file.
        // You can return it, write it to a file, or send it over a network as needed.
    }

    // Example method to provide an input stream.
    // Replace this with your actual implementation for obtaining the Excel data.
    static MemoryStream GetInputStream()
    {
        // Create a simple workbook with a single cell for demonstration.
        Workbook wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("Original Value");

        // Save the workbook to a memory stream.
        MemoryStream ms = new MemoryStream();
        wb.Save(ms, SaveFormat.Xlsx);
        ms.Position = 0; // Reset position for reading.
        return ms;
    }
}
