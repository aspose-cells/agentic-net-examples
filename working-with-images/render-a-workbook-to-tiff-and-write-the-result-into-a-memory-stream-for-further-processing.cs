// Title: How to export an Aspose.Cells Workbook as a multi‑page TIFF into a MemoryStream and then save it to a file in C#
// AI Prompts: Write C# code that loads or creates an Aspose.Cells Workbook and saves it directly to a MemoryStream using SaveFormat.Tiff. | Show how to reset the MemoryStream position and copy its contents to a FileStream to create a .tiff file on disk. | Include robust try‑catch blocks for the workbook‑to‑TIFF conversion and file‑writing steps.
// Common Searches: c# export Aspose.Cells workbook to multi page TIFF in memory stream | aspose.cells save workbook as TIFF without intermediate file | how to write TIFF MemoryStream to file using C# | example of SaveFormat.Tiff with Aspose.Cells and MemoryStream
// Tags: export workbook to multi‑page TIFF Aspose.Cells | save Aspose.Cells workbook as TIFF in memory stream | write TIFF MemoryStream to file C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates or loads an Aspose.Cells Workbook, adds sample data, and uses workbook.Save with SaveFormat.Tiff to write a multi‑page TIFF directly into a MemoryStream. After resetting the stream position, the code copies the stream to a FileStream, creating an output.tiff file, while handling errors with try‑catch blocks.
    public class WorkbookToTiffConverter
    {
        public MemoryStream ConvertToTiff()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Example: add some data to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Text");
                sheet.Cells["B2"].PutValue(12345);

                // Prepare a memory stream to hold the TIFF output
                MemoryStream tiffStream = new MemoryStream();

                // Save the workbook as a multi‑page TIFF into the memory stream
                workbook.Save(tiffStream, SaveFormat.Tiff);

                // Reset the stream position for further processing
                tiffStream.Position = 0;

                return tiffStream;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during TIFF conversion: {ex.Message}");
                throw;
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                WorkbookToTiffConverter converter = new WorkbookToTiffConverter();
                using (MemoryStream tiffStream = converter.ConvertToTiff())
                {
                    // Define output file path
                    string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.tiff");

                    // Ensure the directory exists
                    string? directory = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    // Write the TIFF stream to a file
                    using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        tiffStream.CopyTo(fileStream);
                    }

                    Console.WriteLine($"TIFF file successfully saved to: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
