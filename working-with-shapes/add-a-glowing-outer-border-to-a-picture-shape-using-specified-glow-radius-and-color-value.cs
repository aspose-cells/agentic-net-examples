using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AddGlowToPictureApp
{
    class AddGlowToPicture
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                string imagePath = "sampleImage.jpg";

                // Ensure the image file exists before adding it
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }
                else
                {
                    // Add the picture to the worksheet
                    int pictureIndex = sheet.Pictures.Add(2, 2, imagePath);
                    Picture picture = sheet.Pictures[pictureIndex];

                    // Optional border for visibility
                    picture.BorderLineColor = Color.Black;
                    picture.BorderWeight = 1.0; // 1 pt

                    // Apply glow effect
                    GlowEffect glow = picture.Glow;
                    CellsColor glowColor = workbook.CreateCellsColor();
                    glowColor.Color = Color.FromArgb(255, 255, 0, 0); // Red glow
                    glow.Color = glowColor;
                    glow.Size = 12;            // Glow radius in points
                    glow.Transparency = 0.3;   // 30% transparent
                }

                // Save the workbook
                workbook.Save("PictureWithGlow.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}