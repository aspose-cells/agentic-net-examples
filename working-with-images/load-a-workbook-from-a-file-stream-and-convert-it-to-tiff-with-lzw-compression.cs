// Title: Convert Excel to LZW‑Compressed TIFF from a FileStream with Aspose.Cells for .NET
// Description: Loads an Excel workbook via FileStream, sets ImageOrPrintOptions to TIFF with LZW compression, renders the workbook using WorkbookRender, and writes the result to a TIFF file. Includes checks for source existence and automatic creation of the output folder.
// Keywords: Aspose.Cells TIFF export | LZW compression Excel to TIFF | C# load workbook from stream | WorkbookRender TIFF output | ImageOrPrintOptions LZW | Aspose.Cells file stream conversion | save Excel as compressed TIFF
// Common Searches: Aspose.Cells export Excel to LZW TIFF C# | How to render workbook to TIFF using a FileStream | Convert Excel file to compressed TIFF with Aspose.Cells | C# code for TIFF LZW compression from Excel | Aspose.Cells TIFF rendering options
// Developer Intent: Read an Excel file from a stream and save the entire workbook as a single LZW‑compressed TIFF image.
// Use Cases: Archiving financial spreadsheets as lossless, space‑efficient TIFF files. | Generating printable TIFF previews on a web server without storing the original XLSX. | Batch‑processing incoming Excel streams and delivering compressed TIFFs to a document management system.
// AI Prompts: Generate C# code that uses Aspose.Cells to read an Excel workbook from a FileStream and export it to a single TIFF file with LZW compression. | Explain the role of ImageOrPrintOptions and WorkbookRender when creating LZW‑compressed TIFF images in Aspose.Cells. | Show how to adapt the sample so each worksheet is saved as a separate TIFF page while preserving LZW compression.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing; // For ImageType enum

namespace AsposeCellsExamples
{
    // Loads an Excel workbook via FileStream, sets ImageOrPrintOptions to TIFF with LZW compression, renders the workbook using WorkbookRender, and writes the result to a TIFF file. Includes checks for source existence and automatic creation of the output folder.
    public class WorkbookToTiffLzwDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Path for the output TIFF file
            string tiffPath = "output.tiff";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the workbook from a file stream
                using (FileStream inputStream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
                {
                    Workbook workbook = new Workbook(inputStream);

                    // Configure rendering options for TIFF with LZW compression
                    ImageOrPrintOptions options = new ImageOrPrintOptions
                    {
                        ImageType = ImageType.Tiff,
                        TiffCompression = TiffCompression.CompressionLZW
                    };

                    // Create a renderer for the entire workbook
                    WorkbookRender renderer = new WorkbookRender(workbook, options);

                    // Ensure the output directory exists
                    string outputDir = Path.GetDirectoryName(tiffPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Render the workbook to a TIFF image stream
                    using (FileStream outputStream = new FileStream(tiffPath, FileMode.Create, FileAccess.Write))
                    {
                        renderer.ToImage(outputStream);
                    }
                }

                Console.WriteLine($"Workbook successfully converted to TIFF with LZW compression at: {tiffPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}
