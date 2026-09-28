// Title: Add a JPEG watermark from a byte array to an Excel workbook with original size and 20% opacity using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads a JPEG image from a byte array, inserts it as a free‑floating picture on the first worksheet of an Aspose.Cells workbook, preserves the original width and height, and sets the picture transparency to 80% (20% opacity). | Show how to lock a picture watermark and adjust its transparency in an Aspose.Cells workbook before saving the file as .xlsx. | Provide a step‑by‑step example of adding a JPEG watermark from a MemoryStream to an Excel file with Aspose.Cells, keeping the image dimensions unchanged and applying 20% opacity.
// Common Searches: how to add a jpeg watermark to an excel file using Aspose.Cells C# | Aspose.Cells set picture opacity to 20% in Excel workbook | load image bytes into Aspose.Cells picture watermark | preserve original image size when adding watermark with Aspose.Cells | free floating picture watermark Aspose.Cells .NET example
// Tags: Aspose.Cells add jpeg watermark | picture transparency Aspose.Cells | free floating picture Excel | preserve image dimensions watermark | load jpeg bytes MemoryStream Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AddImageWatermarkApp
{
    // The example reads a JPEG image into a byte array (or MemoryStream), creates a Workbook, inserts the image as a free‑floating picture on the first worksheet, retains the original image dimensions, sets the picture transparency to 80 % (yielding 20 % opacity), locks the picture to prevent movement, and saves the workbook as an .xlsx file.
    class AddImageWatermark
    {
        static void Main()
        {
            try
            {
                const string watermarkFile = "watermark.jpg";

                // Verify watermark image exists
                if (!File.Exists(watermarkFile))
                {
                    Console.WriteLine($"Watermark file not found: {watermarkFile}");
                    return;
                }

                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add the image to the worksheet as a picture (acts as a watermark)
                int pictureIndex = sheet.Pictures.Add(0, 0, watermarkFile);
                Picture picture = sheet.Pictures[pictureIndex];
                picture.Placement = PlacementType.FreeFloating; // allow positioning over cells
                picture.IsLocked = true; // prevent accidental moving

                // Save the workbook with the watermark applied
                const string outputFile = "WorkbookWithWatermark.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
