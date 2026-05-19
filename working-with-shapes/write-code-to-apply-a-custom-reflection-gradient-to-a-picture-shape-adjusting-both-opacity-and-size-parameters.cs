using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    class ApplyCustomReflection
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file to be inserted
                string imagePath = "example.jpg";

                // Verify that the image file exists before adding it
                if (File.Exists(imagePath))
                {
                    // Add the picture to the worksheet
                    int pictureIndex = worksheet.Pictures.Add(0, 0, imagePath);
                    Picture picture = worksheet.Pictures[pictureIndex];

                    // The Picture object derives from Shape, so we can access reflection through it
                    Shape shape = picture;

                    // Configure a custom reflection effect
                    ReflectionEffect reflection = shape.Reflection;
                    reflection.Type = ReflectionEffectType.Custom;   // Enable custom settings
                    reflection.Size = 80;        // End alpha position (percentage)
                    reflection.Transparency = 0.3; // Starting transparency (0 = opaque, 1 = clear)
                    reflection.Blur = 5;         // Blur radius in points (optional)
                    reflection.Distance = 10;    // Distance from the shape in points (optional)
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Save the workbook with the applied reflection effect (if any)
                string outputPath = "CustomReflectionPicture.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}