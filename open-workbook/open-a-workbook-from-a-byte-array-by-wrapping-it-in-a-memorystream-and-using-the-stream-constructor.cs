// Title: Open an Excel workbook from a byte array using Aspose.Cells MemoryStream constructor in C#
// Common Searches: C# Aspose.Cells load workbook from a byte[] using MemoryStream | How to create an Aspose.Cells Workbook from Excel bytes in .NET | Read Excel file stored as byte array into Aspose.Cells workbook | Using Stream constructor to open .xlsx from memory with Aspose.Cells | Load workbook from byte array without saving to disk Aspose.Cells C#
// Tags: load workbook from MemoryStream Aspose.Cells | create workbook from byte array stream C# | Aspose.Cells Workbook constructor Stream overload | open Excel bytes in .NET without file | initialize Aspose.Cells workbook with in‑memory data

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates wrapping an Excel byte array in a MemoryStream and passing it to the Aspose.Cells Workbook constructor that accepts a Stream, allowing the workbook to be opened directly from memory and accessed programmatically.
class Program
{
    static void Main()
    {
        // Obtain the Excel file as a byte array (replace with your actual source)
        byte[] excelBytes = GetExcelFileBytes();

        // Wrap the byte array in a MemoryStream
        using (MemoryStream stream = new MemoryStream(excelBytes))
        {
            // Load the workbook from the stream using the Workbook constructor that accepts a Stream
            Workbook workbook = new Workbook(stream);

            // Example usage: print the name of the first worksheet
            Console.WriteLine("First worksheet: " + workbook.Worksheets[0].Name);
        }
    }

    // Placeholder method to retrieve the byte array; replace with your actual implementation.
    static byte[] GetExcelFileBytes()
    {
        // Example: read an existing .xlsx file into a byte array.
        // In practice, the bytes could come from a database, web service, etc.
        return File.ReadAllBytes("sample.xlsx");
    }
}
