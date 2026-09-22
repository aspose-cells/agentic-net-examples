// Title: Add a full‑page PNG watermark with 40% opacity to an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads a PNG file into a byte array, creates a MemoryStream, inserts the image as a picture on the first worksheet with FreeFloating placement, scales it to the worksheet's printable page dimensions, and sets the picture opacity to 40% using Aspose.Cells. | Show how to compute the printable page width and height in pixels from the worksheet PageSetup and apply those dimensions to a picture object so the watermark covers the entire sheet without moving with cells.
// Common Searches: c# aspose.cells add semi transparent PNG watermark to entire Excel sheet | scale picture to printable area of worksheet aspose.cells c# | set image opacity when inserting picture in aspose.cells workbook | use memory stream to add PNG watermark to excel using aspose.cells | how to make worksheet picture free floating in aspose.cells .net
// Tags: Aspose.Cells add PNG watermark from byte array | worksheet picture scaling to printable area Aspose.Cells | apply 40% transparency to worksheet image Aspose.Cells | non‑moving picture insertion Aspose.Cells | memory stream image insertion Aspose.Cells C#

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example reads a PNG file into a byte array, creates a MemoryStream, and adds the image to the first worksheet as a free‑floating picture. It calculates the printable page size, resizes the picture to cover the whole page, sets the opacity to 40%, and saves the workbook as WatermarkedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string pngPath = "watermark.png";

            // Verify that the watermark image exists
            if (!File.Exists(pngPath))
            {
                Console.WriteLine($"File not found: {pngPath}");
                return;
            }

            // Load PNG image bytes
            byte[] pngBytes = File.ReadAllBytes(pngPath);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add the picture to the worksheet from a memory stream
            using (MemoryStream ms = new MemoryStream(pngBytes))
            {
                int pictureIdx = sheet.Pictures.Add(0, 0, ms);
                Picture picture = sheet.Pictures[pictureIdx];

                // Scale the picture to match the page size (96 DPI)
                double pageWidthInInches = sheet.PageSetup.PaperWidth;
                double pageHeightInInches = sheet.PageSetup.PaperHeight;
                picture.Width = (int)(pageWidthInInches * 96);
                picture.Height = (int)(pageHeightInInches * 96);

                // Optional: set placement so the picture does not move with cells
                picture.Placement = PlacementType.FreeFloating;
            }

            // Save the workbook
            const string outputPath = "WatermarkedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
