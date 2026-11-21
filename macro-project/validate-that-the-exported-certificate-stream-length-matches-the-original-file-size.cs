// Title: How to verify that a workbook saved to a MemoryStream retains the original XLSX file size using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# compare original XLSX file size with MemoryStream length after workbook.Save | check if exported Excel stream size matches source file size in .NET | validate workbook.Save memory stream byte count equals original file length Aspose.Cells | detect size mismatch when saving workbook to MemoryStream using Aspose.Cells | C# get original file size and compare to saved stream length Aspose.Cells
// Tags: workbook.Save memory stream Aspose.Cells | source file size vs exported stream length | Excel byte count validation C# | Aspose.Cells exported stream size check | verify XLSX size consistency after Save

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX file into an Aspose.Cells Workbook, records the original file size, saves the workbook to a MemoryStream with SaveFormat.Xlsx, and then compares the stream's Length to the original size, outputting whether the sizes match.
class Program
{
    static void Main()
    {
        // Path to the original Excel file
        string filePath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(filePath);

        // Get the original file size in bytes
        long originalFileSize = new FileInfo(filePath).Length;

        // Export the workbook to a memory stream in the same format
        using (MemoryStream exportedStream = new MemoryStream())
        {
            workbook.Save(exportedStream, SaveFormat.Xlsx);

            // Length of the exported stream (in bytes)
            long exportedSize = exportedStream.Length;

            // Validate that the exported stream length matches the original file size
            if (exportedSize == originalFileSize)
            {
                Console.WriteLine("Exported stream length matches the original file size.");
            }
            else
            {
                Console.WriteLine($"Size mismatch: original = {originalFileSize} bytes, exported = {exportedSize} bytes.");
            }
        }
    }
}
