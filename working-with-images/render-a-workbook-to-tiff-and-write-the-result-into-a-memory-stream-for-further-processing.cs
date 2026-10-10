// Title: Save an Aspose.Cells Workbook as a TIFF image into a MemoryStream and write it to a file using C#
// AI Prompts: Write C# code that creates an Aspose.Cells Workbook, adds sample data, and saves it directly to a MemoryStream in TIFF format. | Demonstrate how to reset the MemoryStream position and copy its contents to a FileStream to generate a .tiff file on disk. | Add comprehensive try‑catch blocks that log any conversion errors while using Aspose.Cells SaveFormat.Tiff.
// Common Searches: aspnet convert excel workbook to tiff using memory stream | c# Aspose.Cells save workbook as tiff without writing to disk first | how to export Aspose.Cells worksheet to tiff image in memory | write tiff stream from Aspose.Cells to file in C# | sample code for Aspose.Cells SaveFormat.Tiff with MemoryStream
// Tags: Aspose.Cells SaveFormat.Tiff to MemoryStream | C# workbook to TIFF conversion | memory stream TIFF output Aspose.Cells | copy TIFF MemoryStream to FileStream | error handling Aspose.Cells image export

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a new Workbook, inserts sample values, and saves it directly to a MemoryStream in TIFF format using Aspose.Cells. After resetting the stream position, the TIFF data is copied to a FileStream, creating an output.tiff file. The code includes directory creation and robust error handling for the conversion process.
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

                // Save the workbook directly to TIFF format into the memory stream
                workbook.Save(tiffStream, SaveFormat.Tiff);

                // Reset the stream position to the beginning for further processing
                tiffStream.Position = 0;

                // Return the memory stream containing the TIFF image
                return tiffStream;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during conversion: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                WorkbookToTiffConverter converter = new WorkbookToTiffConverter();
                using (MemoryStream tiffStream = converter.ConvertToTiff())
                {
                    // Define output file path
                    string outputPath = "output.tiff";

                    // Ensure the directory exists
                    string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    // Write the TIFF stream to a file
                    using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        tiffStream.CopyTo(fileStream);
                    }

                    Console.WriteLine($"TIFF file successfully saved to: {Path.GetFullPath(outputPath)}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
