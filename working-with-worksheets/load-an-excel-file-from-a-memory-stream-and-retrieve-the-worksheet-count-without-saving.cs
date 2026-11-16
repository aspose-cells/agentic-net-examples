// Title: Load an Excel workbook from a MemoryStream and get the worksheet count using Aspose.Cells for .NET
// Common Searches: how to create an Aspose.Cells Workbook from a byte array in C# | retrieve number of worksheets in an Excel file loaded from MemoryStream with Aspose.Cells | read Excel file bytes and count sheets without saving using Aspose.Cells .NET | C# load Excel workbook from stream and get sheet count with Aspose.Cells | Aspose.Cells get worksheet count without persisting workbook
// Tags: load workbook from memory stream Aspose.Cells | retrieve worksheet count Aspose.Cells | read Excel bytes without saving Aspose.Cells | in-memory Excel processing C# Aspose.Cells | worksheet count property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example shows how to instantiate a MemoryStream from a byte array containing Excel data, create an Aspose.Cells Workbook from that stream, access the Worksheets.Count property to obtain the number of worksheets, and output the count—all without writing the file back to disk.
class Program
{
    static void Main()
    {
        // Assume you have the Excel file bytes in a byte array called excelBytes
        byte[] excelBytes = GetExcelFileBytes(); // Replace with your actual source

        // Load the Excel file from a memory stream
        using (MemoryStream ms = new MemoryStream(excelBytes))
        {
            // Create a Workbook instance from the memory stream
            Workbook workbook = new Workbook(ms);

            // Retrieve the number of worksheets in the workbook
            int worksheetCount = workbook.Worksheets.Count;

            // Output the count (or use it as needed)
            Console.WriteLine($"Worksheet count: {worksheetCount}");
        }
    }

    // Placeholder method to obtain Excel file bytes.
    // Replace this with your actual implementation (e.g., reading from a file, database, etc.).
    static byte[] GetExcelFileBytes()
    {
        // Example: read from a file on disk (adjust the path as needed)
        // return File.ReadAllBytes(@"C:\Path\To\Your\File.xlsx");

        // For demonstration, return an empty array.
        return new byte[0];
    }
}
