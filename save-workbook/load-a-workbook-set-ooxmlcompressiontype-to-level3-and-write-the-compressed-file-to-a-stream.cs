using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook from a file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Create OOXML save options and set compression to Level3
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.CompressionType = OoxmlCompressionType.Level3;

        // Save the workbook to a memory stream using the specified options
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, saveOptions);

            // Reset the stream position if further processing is needed
            stream.Position = 0;

            // Example: write the compressed stream to a file for verification
            using (FileStream file = new FileStream("output_compressed.xlsx", FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(file);
            }
        }
    }
}