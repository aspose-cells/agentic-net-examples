// Title: Generate a PNG thumbnail of the first worksheet in an Excel file using Aspose.Cells for .NET
// Common Searches: C# code to create a thumbnail image of the first sheet in an Excel workbook with Aspose.Cells | How to render the first worksheet to a PNG preview using Aspose.Cells .NET | Aspose.Cells ImageOrPrintOptions settings for 150 DPI thumbnail generation | Save first worksheet as PNG file with SheetRender in C# | Create Excel sheet preview image programmatically with Aspose.Cells
// Tags: Aspose.Cells PNG sheet rendering | 150 DPI thumbnail generation Aspose | C# export worksheet as image | Excel sheet preview creation Aspose.Cells | first worksheet thumbnail generation .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsThumbnail
{
    // The example checks for an input.xlsx file, loads it into an Aspose.Cells Workbook, configures ImageOrPrintOptions for a 150 DPI PNG thumbnail with one page per sheet, uses SheetRender to render the first worksheet to thumbnail.png (creating the output folder if needed), and writes the thumbnail path to the console.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure image rendering options for a PNG thumbnail
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Default image format is PNG, so no need to set explicitly
                    OnePagePerSheet = true,
                    HorizontalResolution = 150,
                    VerticalResolution = 150
                };

                // Render the first worksheet to the thumbnail image
                SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], imgOptions);
                string thumbnailPath = "thumbnail.png";

                // Ensure the directory for the thumbnail exists
                string thumbnailDir = Path.GetDirectoryName(thumbnailPath);
                if (!string.IsNullOrEmpty(thumbnailDir) && !Directory.Exists(thumbnailDir))
                {
                    Directory.CreateDirectory(thumbnailDir);
                }

                sheetRender.ToImage(0, thumbnailPath);

                Console.WriteLine($"Thumbnail created at: {thumbnailPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
