// Title: Validate that a workbook saved to a MemoryStream with Aspose.Cells .NET has the same byte size as the original XLSX file
// AI Prompts: Write C# code that loads an existing .xlsx file using Aspose.Cells, saves the workbook to a MemoryStream with SaveFormat.Xlsx, and prints whether the stream length equals the original file size. | Create a method that takes a file path, exports the workbook to a MemoryStream via Aspose.Cells, compares MemoryStream.Length to the source file length, and returns a boolean indicating a size match.
// Common Searches: how to verify that Aspose.Cells saved workbook size matches original file in C# | C# compare original Excel file length with MemoryStream length after SaveFormat.Xlsx | Aspose.Cells .NET check for size mismatch when exporting to memory stream | determine if exported XLSX stream is identical in size to source file using Aspose.Cells | validate workbook byte size after saving to MemoryStream with Aspose.Cells
// Tags: Aspose.Cells compare exported stream size | C# memory stream length verification for XLSX | SaveFormat.Xlsx size check with Aspose.Cells | validate workbook byte size after save | detect Excel file size mismatch in .NET

using System;
using System.IO;
using Aspose.Cells;

// The program loads "input.xlsx", records its original byte size, saves the workbook to a MemoryStream in XLSX format using Aspose.Cells, compares the stream length to the original size, and outputs whether the sizes match.
class Program
{
    static void Main()
    {
        // Path to the original Excel file
        string filePath = "input.xlsx";

        // Get the original file size in bytes
        long originalSize = new FileInfo(filePath).Length;

        // Load the workbook from the file
        Workbook workbook = new Workbook(filePath);

        // Export the workbook to a memory stream using the same format
        using (MemoryStream exportedStream = new MemoryStream())
        {
            workbook.Save(exportedStream, SaveFormat.Xlsx);
            long exportedSize = exportedStream.Length;

            // Validate that the exported stream length matches the original file size
            if (exportedSize == originalSize)
            {
                Console.WriteLine("Exported stream length matches the original file size.");
            }
            else
            {
                Console.WriteLine($"Size mismatch: original = {originalSize} bytes, exported = {exportedSize} bytes.");
            }
        }
    }
}
