// Title: Save an Excel workbook with OOXML Level 3 compression to a MemoryStream using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an existing .xlsx file, configures OoxmlSaveOptions to use Level3 compression, and saves the workbook into a MemoryStream with Aspose.Cells. | Show how to copy the compressed workbook from a MemoryStream to a physical file after saving with OoxmlCompressionType.Level3 in .NET.
// Common Searches: Aspose.Cells how to set OoxmlCompressionType Level3 when saving a workbook | C# save Excel file to MemoryStream with OOXML compression using Aspose.Cells | Compress XLSX output to Level3 using Aspose.Cells .NET API | Write compressed workbook stream to file after using OoxmlSaveOptions
// Tags: OoxmlSaveOptions Level3 compression | save workbook to MemoryStream Aspose.Cells | compress XLSX with Aspose.Cells .NET | copy compressed stream to file C# | Aspose.Cells workbook compression settings

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing .xlsx workbook, sets OoxmlSaveOptions.CompressionType to Level3, saves the workbook into a MemoryStream, resets the stream position, and then copies the compressed data to a new file.
class Program
{
    static void Main()
    {
        // Load an existing workbook from a file
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Create OOXML save options and set compression to Level3
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.CompressionType = OoxmlCompressionType.Level3;

        // Save the workbook to a memory stream using the specified options
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, saveOptions);

            // Reset stream position if further processing is needed
            stream.Position = 0;

            // Example: write the compressed stream to a file
            using (FileStream file = new FileStream("output_compressed.xlsx", FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(file);
            }
        }
    }
}
