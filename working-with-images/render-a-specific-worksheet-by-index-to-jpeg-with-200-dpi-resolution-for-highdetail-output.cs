// Title: Render a specific worksheet to high‑resolution JPEG (200 DPI) using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells render worksheet at index to 200 DPI JPEG | How to export each page of an Excel sheet as separate JPEG files with Aspose.Cells | Command‑line tool to convert a selected worksheet to high‑resolution JPEG using Aspose.Cells | Set ImageOrPrintOptions DPI for JPEG output in Aspose.Cells C# | Render multiple pages of a worksheet to JPEG with Aspose.Cells .NET
// Tags: SheetRender export worksheet to JPEG | ImageOrPrintOptions 200 DPI JPEG | Aspose.Cells render specific sheet by index | C# generate multi‑page JPEG from Excel sheet | command‑line worksheet selection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsRenderDemo
{
    // The example loads an Excel workbook, validates a worksheet index passed via the command line, configures ImageOrPrintOptions for JPEG format at 200 DPI, and uses SheetRender to create a separate high‑resolution JPEG file for each page of the selected worksheet.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 1)
            {
                Console.WriteLine("Error: Input file path is required.");
                Console.WriteLine("Usage: AsposeCellsRenderDemo.exe <inputFilePath> [worksheetIndex]");
                return;
            }

            string inputPath = args[0];
            int sheetIndex = 0;

            if (args.Length >= 2 && !int.TryParse(args[1], out sheetIndex))
            {
                Console.WriteLine("Error: Worksheet index must be a valid integer.");
                return;
            }

            // Verify the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Validate worksheet index
                if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
                {
                    Console.WriteLine($"Error: Worksheet index {sheetIndex} is out of range. The workbook contains {workbook.Worksheets.Count} worksheet(s).");
                    return;
                }

                Worksheet worksheet = workbook.Worksheets[sheetIndex];
                Console.WriteLine($"Rendering worksheet \"{worksheet.Name}\" (index {sheetIndex}) from \"{Path.GetFileName(inputPath)}\".");

                // Configure image rendering options
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Jpeg,          // Output format JPEG
                    HorizontalResolution = 200,          // 200 DPI horizontal
                    VerticalResolution = 200,            // 200 DPI vertical
                    OnePagePerSheet = false              // Keep original pagination
                };

                // Create the sheet renderer (SheetRender does not implement IDisposable)
                SheetRender sheetRender = new SheetRender(worksheet, options);
                try
                {
                    int pageCount = sheetRender.PageCount;
                    Console.WriteLine($"Total pages to render: {pageCount}");

                    // Render each page to a separate JPEG file
                    for (int page = 0; page < pageCount; page++)
                    {
                        try
                        {
                            string outputFileName = $"{Path.GetFileNameWithoutExtension(inputPath)}_Sheet{sheetIndex}_Page{page + 1}.jpg";
                            sheetRender.ToImage(page, outputFileName);
                            Console.WriteLine($"Page {page + 1} saved as \"{outputFileName}\".");
                        }
                        catch (Exception exPage)
                        {
                            Console.WriteLine($"Error rendering page {page + 1}: {exPage.Message}");
                        }
                    }
                }
                finally
                {
                    // Explicitly release resources used by SheetRender
                    sheetRender.Dispose();
                }

                Console.WriteLine("Rendering completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
