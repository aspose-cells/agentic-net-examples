// Title: Create a PNG thumbnail of the first page of a multi‑page TIFF workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a multi‑page TIFF workbook with Aspose.Cells, configures ImageOrPrintOptions for a low‑resolution PNG, and renders only the first worksheet to a thumbnail file. | Show how to verify the input TIFF exists, use SheetRender to export the first sheet as a PNG image, and handle exceptions in a .NET console application.
// Common Searches: asp.net how to export first sheet of a tiff workbook to png thumbnail using Aspose.Cells | c# create low‑resolution preview image of first page in a multi‑page tiff Excel file | sample code for rendering first worksheet of a tiff workbook to a png file with Aspose.Cells | generate thumbnail from tiff workbook page with Aspose.Cells ImageOrPrintOptions | convert first page of tiff workbook to png using sheetrender c#
// Tags: Aspose.Cells render first worksheet to PNG | C# ImageOrPrintOptions thumbnail generation | SheetRender export TIFF workbook page | multi-page TIFF workbook preview Aspose.Cells | low-resolution PNG thumbnail Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace ThumbnailGenerator
{
    // // Loads a TIFF workbook, checks that it contains worksheets, sets 96 DPI ImageOrPrintOptions, uses SheetRender to render the first worksheet as a PNG thumbnail, saves the image to the specified path, and handles errors gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input TIFF workbook and output thumbnail image
            string inputPath = "input.tif";
            string outputPath = "thumbnail.png";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the TIFF workbook (first page corresponds to the first worksheet)
                Workbook workbook = new Workbook(inputPath);

                // Ensure there is at least one worksheet to render
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook contains no worksheets.");
                    return;
                }

                // Configure image options for a PNG thumbnail
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Horizontal and vertical resolution (DPI) for the thumbnail
                    HorizontalResolution = 96,
                    VerticalResolution = 96,
                    OnePagePerSheet = true // Ensure one page per sheet
                };

                // Render only the first worksheet (first page) to an image
                SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], imgOptions);
                sheetRender.ToImage(0, outputPath);

                Console.WriteLine($"Thumbnail saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
