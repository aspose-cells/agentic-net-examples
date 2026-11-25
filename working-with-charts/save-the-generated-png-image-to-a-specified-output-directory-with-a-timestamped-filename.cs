// Title: Save each rendered worksheet page as a timestamped PNG file to a chosen folder using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet pages to PNG with unique timestamped filenames | How to render Excel sheet to PNG and save to a specific directory using Aspose.Cells | Aspose.Cells generate PNG images for each page of a workbook and include date‑time in the file name | Create filesystem‑safe file names for Excel image export in C# with Aspose.Cells | Configure ImageOrPrintOptions for PNG output and store images in a custom folder using Aspose.Cells
// Tags: Aspose.Cells worksheet page to PNG rendering | C# timestamped file naming for Excel image export | ImageOrPrintOptions PNG configuration Aspose.Cells | SheetRender save page as image Aspose.Cells | Create output directory and safe filenames Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing; // Required for ImageType enum

namespace AsposeCellsImageExport
{
    // The sample loads an Excel workbook, fits the first worksheet to a single page, and uses Aspose.Cells' ImageOrPrintOptions (PNG) together with SheetRender to render each page. It ensures the output folder exists, builds a filesystem‑safe filename that includes the worksheet name, page number and a high‑resolution timestamp, and saves each rendered page as a PNG file in the specified directory.
    class Program
    {
        static int Main(string[] args)
        {
            // Default input and output locations
            string inputFilePath = "input.xlsx";
            string outputDirectory = "output";

            // Override via command‑line arguments
            if (args.Length > 0)
                inputFilePath = args[0];
            if (args.Length > 1)
                outputDirectory = args[1];

            // Verify that the input file exists
            if (!File.Exists(inputFilePath))
            {
                Console.Error.WriteLine($"Error: Input file \"{inputFilePath}\" does not exist.");
                return 1;
            }

            try
            {
                // Ensure output directory exists
                Directory.CreateDirectory(outputDirectory);

                // Load workbook
                Workbook workbook = new Workbook(inputFilePath);
                Worksheet worksheet = workbook.Worksheets[0]; // First worksheet

                // Fit worksheet to a single page
                worksheet.PageSetup.SetFitToPages(1, 1);
                worksheet.PageSetup.PrintArea = worksheet.Cells.MaxDisplayRange.RefersTo;

                // Set rendering options for PNG output
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png, // Use the correct enum from Aspose.Cells.Drawing
                    Transparent = false,
                    HorizontalResolution = 96,
                    VerticalResolution = 96
                };

                // Create renderer
                SheetRender sheetRender = new SheetRender(worksheet, renderOptions);

                // Render each page
                for (int pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
                {
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
                    string safeSheetName = MakeFileSystemSafe(worksheet.Name);
                    string fileName = $"Sheet_{safeSheetName}_Page_{pageIndex + 1}_{timestamp}.png";
                    string fullPath = Path.Combine(outputDirectory, fileName);

                    // Save page as image file
                    sheetRender.ToImage(pageIndex, fullPath);
                    Console.WriteLine($"Page {pageIndex + 1} saved as \"{fullPath}\"");
                }

                Console.WriteLine("Image export completed successfully.");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Exception: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
                return 1;
            }
        }

        // Replace invalid filename characters with underscores
        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
