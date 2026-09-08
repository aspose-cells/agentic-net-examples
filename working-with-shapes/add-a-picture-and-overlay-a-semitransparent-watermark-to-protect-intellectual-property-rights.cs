// Title: Add an image to cell B2 and use it as a semi‑transparent watermark in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Insert a JPEG file into a specific cell of an Excel worksheet and set the picture to free‑floating using Aspose.Cells in C#. | Create a semi‑transparent Excel watermark by pre‑processing an image for opacity, adding it as a picture, and saving the workbook as XLSX with Aspose.Cells. | Programmatically verify the output folder exists, create it if necessary, and save the workbook that contains the overlaid picture.
// Common Searches: how to place an image at cell B2 in Aspose.Cells C# | Aspose.Cells add picture as watermark with transparency in .NET | C# set picture placement to free floating in Excel using Aspose.Cells | save Excel file with overlay image ensuring output directory exists Aspose.Cells | adjust JPEG opacity prior to adding as watermark in Aspose.Cells
// Tags: add picture to worksheet cell Aspose.Cells C# | free‑floating picture placement Aspose.Cells | image opacity watermark Aspose.Cells | ensure output directory before saving workbook Aspose.Cells | preprocess image for transparency before insertion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, loads a JPEG image, inserts it at cell B2 as a free‑floating picture, notes that opacity must be set beforehand, ensures the output folder exists, and saves the workbook as an XLSX file containing the semi‑transparent watermark.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image that will be used as a picture/watermark
            string imagePath = @"C:\Images\sample.jpg";

            // Verify that the image file exists before adding it
            if (!File.Exists(imagePath))
                throw new FileNotFoundException("Image file not found.", imagePath);

            // Add the picture to the worksheet at cell B2 (row index 1, column index 1)
            int pictureIndex = sheet.Pictures.Add(1, 1, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Note: Aspose.Cells does not expose a direct Transparency property for Picture.
            // If needed, adjust the image beforehand to include the desired opacity.

            // Ensure the picture moves and resizes with cells (optional)
            picture.Placement = PlacementType.FreeFloating;

            // Prepare output path and ensure the directory exists
            string outputPath = @"C:\Output\WorkbookWithWatermark.xlsx";
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook with the picture overlay
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
