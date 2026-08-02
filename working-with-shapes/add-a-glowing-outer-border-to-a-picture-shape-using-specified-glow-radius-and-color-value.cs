// Title: Apply Glow Effect and Outer Border to a Picture Shape in Aspise.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a JPEG picture, optionally adds a black border, then sets a glow effect with a configurable radius, color, and transparency before saving the file as an XLSX document.
// Keywords: Aspose.Cells picture glow | C# picture border | add outer glow Aspose.Cells | picture shape effect .NET | glow radius color transparency | Excel workbook image styling | Aspose.Cells example C#
// Common Searches: Aspose.Cells add glow to picture | C# set picture border and glow in Excel | how to apply outer glow effect to image in Aspose.Cells | configure picture glow radius .NET | sample code picture glow Aspose.Cells
// Developer Intent: Add a colored glow with a defined radius around a picture shape and optionally draw a visible border using Aspose.Cells for .NET.
// Use Cases: Insert a JPEG into a worksheet, apply a 2‑point black border, and add a yellow glow of 12 points with 30% transparency. | Validate image file existence before adding it to avoid runtime errors. | Reuse the same glow configuration for multiple pictures by applying identical GlowEffect settings.
// AI Prompts: Generate C# code that inserts an image into an Aspose.Cells worksheet, applies a red glow of 15 points with 20% transparency, and saves the workbook. | Create a reusable method in C# for Aspose.Cells that adds a picture with customizable border weight, border color, glow radius, glow color, and glow transparency.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Example
{
    // Creates a new workbook, inserts a JPEG picture, optionally adds a black border, then sets a glow effect with a configurable radius, color, and transparency before saving the file as an XLSX document.
    class AddGlowToPicture
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                string imagePath = "sample.jpg";

                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }
                else
                {
                    // Add picture to the worksheet
                    int picIndex = sheet.Pictures.Add(2, 2, imagePath);
                    Picture picture = sheet.Pictures[picIndex];

                    // Optional: set a visible border around the picture
                    picture.BorderLineColor = Color.Black;
                    picture.BorderWeight = 2; // border weight in points

                    // Configure the glow effect for the picture
                    GlowEffect glow = picture.Glow;
                    glow.Size = 12; // glow radius in points
                    CellsColor glowColor = workbook.CreateCellsColor();
                    glowColor.Color = Color.Yellow; // desired glow color
                    glow.Color = glowColor;
                    glow.Transparency = 0.3; // 30% transparency (optional)
                }

                // Save the workbook
                string outputPath = "PictureWithGlow.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
