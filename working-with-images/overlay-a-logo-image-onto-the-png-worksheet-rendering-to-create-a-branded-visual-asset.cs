// Title: Overlay a Logo on a PNG Render of an Excel Worksheet with Aspose.Cells for .NET (C#)
// Description: C# code that builds a workbook, places a logo picture as a free‑floating shape, renders the first worksheet to a PNG image, and saves the output as a branded file.
// Keywords: Aspose.Cells C# add logo | Excel worksheet to PNG with image | overlay picture on rendered sheet | free floating picture Aspose.Cells | .NET brand Excel image | render worksheet as PNG Aspose
// Common Searches: How to add a logo to a PNG image generated from an Excel sheet using Aspose.Cells | C# render Excel worksheet to PNG with a watermark image | Insert picture into worksheet before exporting to PNG in Aspose.Cells | Create branded Excel snapshot as PNG in .NET | Place free‑floating image on Excel sheet for rendering
// Developer Intent: Insert a logo into a worksheet and export the sheet as a branded PNG image.
// Use Cases: Automated generation of company‑branded report thumbnails for email newsletters | Creating watermarked screenshots of spreadsheets for documentation or compliance | Producing preview images with a logo for a web gallery of Excel‑based dashboards
// AI Prompts: Show C# code to position a logo at the top‑right corner with a 20‑pixel offset before rendering the worksheet to PNG using Aspose.Cells. | Generate a multi‑sheet workbook where each page is rendered to PNG with a semi‑transparent watermark applied via Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsBrandingDemo
{
    // C# code that builds a workbook, places a logo picture as a free‑floating shape, renders the first worksheet to a PNG image, and saves the output as a branded file.
    class Program
    {
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Prepare a workbook (you can also load an existing file)
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add some sample data so the rendered sheet is not empty
                sheet.Cells["A1"].PutValue("Branding Demo");
                sheet.Cells["A2"].PutValue(DateTime.Now.ToString());

                // ------------------------------------------------------------
                // 2. Define paths for the rendered image and the logo
                // ------------------------------------------------------------
                string sheetImagePath = "sheet_render.png";
                string logoPath = "logo.png";
                string brandedImagePath = "branded_output.png";

                // ------------------------------------------------------------
                // 3. Set rendering options
                // ------------------------------------------------------------
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    OnePagePerSheet = true
                };

                // ------------------------------------------------------------
                // 4. Add logo as a picture shape if the file exists
                // ------------------------------------------------------------
                if (File.Exists(logoPath))
                {
                    try
                    {
                        // Add picture at the top‑left cell (row 0, column 0)
                        int pictureIndex = sheet.Pictures.Add(0, 0, logoPath);
                        Picture picture = sheet.Pictures[pictureIndex];
                        picture.Placement = PlacementType.FreeFloating;

                        // Optional: offset the picture slightly from the top‑left corner
                        const int margin = 10;
                        picture.Left = margin;
                        picture.Top = margin;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Unable to add logo. {ex.Message}");
                    }
                }

                // ------------------------------------------------------------
                // 5. Render the worksheet to a PNG image file
                // ------------------------------------------------------------
                try
                {
                    SheetRender sheetRender = new SheetRender(sheet, renderOptions);
                    sheetRender.ToImage(0, sheetImagePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during rendering: {ex.Message}");
                }

                // ------------------------------------------------------------
                // 6. Rename the rendered file to the final branded image name
                // ------------------------------------------------------------
                if (File.Exists(sheetImagePath))
                {
                    // If a branded image already exists, delete it first
                    if (File.Exists(brandedImagePath))
                    {
                        File.Delete(brandedImagePath);
                    }

                    File.Move(sheetImagePath, brandedImagePath);
                    Console.WriteLine($"Branded image saved to: {brandedImagePath}");
                }
                else
                {
                    Console.WriteLine("Failed to create the rendered sheet image.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
