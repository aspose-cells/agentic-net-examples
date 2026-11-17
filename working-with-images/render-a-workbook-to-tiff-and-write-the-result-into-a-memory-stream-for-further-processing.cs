// Title: Save an Aspose.Cells Workbook as a multi‑page TIFF directly to a MemoryStream in C#
// Common Searches: C# Aspose.Cells export Excel worksheet to TIFF in memory | How to write Aspose.Cells workbook to a MemoryStream as TIFF without creating a file | Aspose.Cells SaveFormat.Tiff to stream example in .NET | Convert Excel workbook to multi-page TIFF using Aspose.Cells and capture output in MemoryStream | Aspose.Cells C# render workbook to TIFF image for further processing
// Tags: Aspose.Cells SaveFormat.Tiff to MemoryStream | C# export workbook to multi-page TIFF | Aspose.Cells render Excel as TIFF image | in-memory TIFF generation with Aspose.Cells | convert Excel to TIFF without disk I/O

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to use Aspose.Cells in C# to save a Workbook as a multi‑page TIFF image directly into a MemoryStream via SaveFormat.Tiff, allowing the image data to be processed further without writing to disk.
public class WorkbookToTiffConverter
{
    /// <returns>MemoryStream containing the TIFF image.</returns>
    public MemoryStream ConvertWorkbookToTiff()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");

            // Prepare a memory stream to hold the TIFF output
            MemoryStream tiffStream = new MemoryStream();

            // Save the workbook as a multi‑page TIFF image into the memory stream
            workbook.Save(tiffStream, SaveFormat.Tiff);

            // Reset the stream position to the beginning for further processing
            tiffStream.Position = 0;

            return tiffStream;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error during workbook to TIFF conversion: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            WorkbookToTiffConverter converter = new WorkbookToTiffConverter();
            using (MemoryStream tiffStream = converter.ConvertWorkbookToTiff())
            {
                const string outputPath = "output.tiff";

                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Write the TIFF data to a file
                using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    tiffStream.CopyTo(fileStream);
                }

                Console.WriteLine($"TIFF file successfully created at: {Path.GetFullPath(outputPath)}");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
