// Title: Apply a Green Outer Glow to a Picture Shape in Excel with Aspose.Cells for .NET
// Description: Creates a workbook, inserts a PNG at cell C3, adds an optional black border, and configures the picture's GlowEffect with a green color, 12‑point radius, and 30 % transparency before saving as PictureWithGlow.xlsx.
// Keywords: Aspose.Cells picture glow | C# Excel outer glow | picture border Aspose.Cells | GlowEffect size color | add glow to worksheet image | Aspose.Cells .NET example
// Common Searches: Aspose.Cells add outer glow to picture | C# set picture glow radius Excel | how to apply green glow to image in Aspose.Cells | configure picture border and glow in .NET | Excel picture glow effect code sample
// Developer Intent: Add a colored outer glow with a defined radius to a picture shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Insert a PNG into a worksheet and highlight it with a green outer glow for visual emphasis. | Combine a thin black border with a customizable glow to match corporate branding. | Generate a placeholder PNG when the source image is missing, then apply identical glow settings.
// AI Prompts: Generate C# code that adds a picture to an Excel sheet with Aspose.Cells and applies a red outer glow of 8 points at 50 % transparency. | Show how to change an existing picture's glow color to blue and increase the radius to 15 points using Aspose.Cells for .NET. | Explain how to programmatically create a placeholder PNG and then set a configurable glow effect on the picture in an Excel workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Example
{
    // Creates a workbook, inserts a PNG at cell C3, adds an optional black border, and configures the picture's GlowEffect with a green color, 12‑point radius, and 30 % transparency before saving as PictureWithGlow.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Ensure the image file exists; create a simple placeholder if missing
                string imagePath = "sample.png";
                if (!File.Exists(imagePath))
                {
                    // 1x1 pixel transparent PNG
                    byte[] pngBytes = Convert.FromBase64String(
                        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2V8AAAAASUVORK5CYII=");
                    File.WriteAllBytes(imagePath, pngBytes);
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add picture to worksheet (row, column, image file path)
                int pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Optional: set a visible border line around the picture
                picture.BorderLineColor = System.Drawing.Color.Black;
                picture.BorderWeight = 2f; // border weight in points

                // Configure glow effect
                GlowEffect glow = picture.Glow;
                CellsColor glowColor = workbook.CreateCellsColor();
                glowColor.Color = System.Drawing.Color.FromArgb(255, 0, 255, 0); // green glow
                glow.Color = glowColor;
                glow.Size = 12; // 12‑point radius
                glow.Transparency = 0.3; // 0.0 = opaque, 1.0 = fully transparent

                // Save the workbook with the picture that has a glowing outer border
                workbook.Save("PictureWithGlow.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
