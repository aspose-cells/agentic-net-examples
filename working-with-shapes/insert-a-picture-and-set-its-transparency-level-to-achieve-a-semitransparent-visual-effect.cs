// Title: Insert a picture into an Excel worksheet and apply 50% fill transparency using Aspose.Cells for .NET (C#)
// Common Searches: c# add image to excel cell with Aspose.Cells and set transparency | Aspose.Cells picture fill transparency example in .NET | render worksheet to png with transparent background using Aspose.Cells C# | how to make an inserted picture semi‑transparent in Aspose.Cells | Aspose.Cells set picture opacity when inserting into Excel
// Tags: insert picture Aspose.Cells C# | picture fill transparency Aspose.Cells | render worksheet to PNG transparent background Aspose.Cells | semi transparent image Excel Aspose.Cells | add image to cell B2 Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace InsertPictureWithTransparency
{
    // The sample creates a placeholder PNG if needed, opens a new workbook, inserts the image at cell B2, resizes it, sets a 50 % fill transparency, adds alternative text, saves the file as XLSX, and then renders the worksheet to a PNG with a transparent background.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the image that will be inserted.
                string imagePath = "sample.png";

                // If the image does not exist, create a simple 1x1 transparent PNG.
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file \"{imagePath}\" not found. Creating a placeholder image.");

                    // Base64 for a 1x1 transparent PNG.
                    const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
                    byte[] pngBytes = Convert.FromBase64String(base64Png);
                    File.WriteAllBytes(imagePath, pngBytes);

                    Console.WriteLine($"Placeholder image saved as \"{imagePath}\".");
                }

                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Insert the picture at cell B2 (row index 1, column index 1).
                int pictureIndex = sheet.Pictures.Add(1, 1, imagePath);
                Picture picture = sheet.Pictures[pictureIndex];

                // Set the picture's size (optional).
                picture.Width = 200;   // pixels
                picture.Height = 200;  // pixels

                // Apply semi‑transparent effect (0 = opaque, 1 = fully transparent).
                picture.Fill.Transparency = 0.5f; // 50 % transparent

                // Optional caption.
                picture.AlternativeText = "Semi‑transparent picture";

                // Save the workbook.
                string outputPath = "PictureWithTransparency.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");

                // -----------------------------------------------------------------
                // Render the worksheet to an image with a transparent background.
                // -----------------------------------------------------------------
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    Transparent = true   // Makes the background of the rendered image transparent.
                };

                SheetRender sheetRender = new SheetRender(sheet, renderOptions);
                sheetRender.ToImage(0, "RenderedSheet.png");
                Console.WriteLine("Worksheet rendered to \"RenderedSheet.png\" with a transparent background.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
