// Title: Load an Excel workbook from a byte[] using Aspose.Cells Workbook(Stream) constructor in C#
// Common Searches: aspnet load excel from byte array without saving file using Aspose.Cells | c# create Workbook from MemoryStream with Aspose.Cells | how to open .xlsx from byte[] in .NET Aspose.Cells | read Excel data from byte[] stream using Aspose.Cells Workbook constructor | load workbook from in‑memory byte array Aspose.Cells C# example
// Tags: Aspose.Cells Workbook constructor Stream overload | C# load Excel workbook from byte array | MemoryStream Excel file loading Aspose.Cells | in‑memory .xlsx loading with Aspose.Cells | byte[] to Workbook conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to wrap a byte array containing an Excel file in a MemoryStream and instantiate an Aspose.Cells Workbook using the Stream constructor, enabling in‑memory loading of .xlsx files without writing to disk.
public class WorkbookLoader
{
    /// <param name="excelBytes">Byte array containing the Excel file.</param>
    /// <returns>Loaded Workbook instance.</returns>
    public static Workbook LoadFromBytes(byte[] excelBytes)
    {
        // Wrap the byte array in a MemoryStream.
        using (MemoryStream stream = new MemoryStream(excelBytes))
        {
            // Use the Workbook constructor that accepts a Stream.
            Workbook workbook = new Workbook(stream);
            return workbook;
        }
    }

    // Example usage
    public static void Main()
    {
        // Example: read a file into a byte array (replace with your source).
        byte[] fileData = File.ReadAllBytes("sample.xlsx");

        // Load the workbook from the byte array.
        Workbook wb = LoadFromBytes(fileData);

        // Perform any workbook operations here.
        Console.WriteLine("Workbook loaded. Worksheets count: " + wb.Worksheets.Count);
    }
}
