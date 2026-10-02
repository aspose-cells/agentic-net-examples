// Title: Save an Excel workbook with OOXML Level‑3 compression to a MemoryStream and write it to a file using Aspose.Cells for .NET
// AI Prompts: Load a workbook, configure OoxmlSaveOptions with CompressionType Level3, and save it directly into a MemoryStream using Aspose.Cells. | Transfer the compressed MemoryStream to a FileStream to generate a Level‑3 compressed XLSX file on disk. | Check for the source file, create the destination folder if missing, and handle exceptions while saving the compressed workbook.
// Common Searches: how to compress an XLSX file with OOXML Level3 using Aspose.Cells in C# | Aspose.Cells save workbook to MemoryStream with compression options | C# example for OoxmlCompressionType.Level3 when exporting Excel | write compressed workbook stream to a file with Aspose.Cells .NET | reduce size of generated XLSX by setting OoxmlSaveOptions compression
// Tags: OOXML Level3 compression Aspose.Cells | memory stream export Aspose.Cells | OoxmlSaveOptions compression C# | compressed XLSX output .NET | ensure output directory exists C#

using System;
using System.IO;
using Aspose.Cells;

// The code loads an existing XLSX file, sets OoxmlSaveOptions.CompressionType to Level3, saves the workbook into a MemoryStream, and then copies the compressed stream to a new file, creating the output folder if necessary and handling errors.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(sourcePath);

            // Configure OOXML compression via save options (Level3)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                CompressionType = OoxmlCompressionType.Level3
            };

            // Save the workbook to a memory stream with the specified compression
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Save(stream, saveOptions);

                // Reset the stream position if you need to read from it later
                stream.Position = 0;

                // Write the compressed stream to an output file
                string outputPath = "output_compressed.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    stream.CopyTo(file);
                }

                Console.WriteLine($"Compressed workbook saved to: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
