// Title: Compress an Excel workbook with OOXML Level 3 compression and save it to a MemoryStream using Aspose.Cells for .NET
// Common Searches: how to set OoxmlCompressionType to Level3 when saving a workbook with Aspose.Cells C# | save an XLSX file to a MemoryStream with maximum compression using Aspose.Cells | Aspose.Cells example for compressing Excel output to Level3 and writing to a stream | C# code to configure OoxmlSaveOptions compression level in Aspose.Cells | write compressed Excel workbook to a byte array using Aspose.Cells
// Tags: OoxmlSaveOptions Level3 compression | save workbook to MemoryStream Aspose.Cells | compress XLSX output Aspose.Cells | set OoxmlCompressionType C# | Aspose.Cells Excel compression options

using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX workbook, configures OoxmlSaveOptions to use Level 3 OOXML compression, and saves the workbook into a MemoryStream, yielding a highly compressed Excel file in memory.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your source as needed)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure OoxmlSaveOptions to use Level3 compression
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
        saveOptions.CompressionType = OoxmlCompressionType.Level3;

        // Write the compressed workbook to a memory stream
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, saveOptions);

            // The stream now contains the compressed XLSX file.
            // Reset the position if you need to read from the beginning.
            stream.Position = 0;

            // Example: write the stream to a file (optional)
            // File.WriteAllBytes("output_compressed.xlsx", stream.ToArray());
        }
    }
}
