// Title: C# – Apply a Custom Reflection Gradient to a Picture Shape with Aspose.Cells
// Description: Creates a workbook, inserts a JPEG picture, and configures Aspose.Cells' ReflectionEffect (Custom) – size, opacity, blur, distance, direction and fade direction – before saving the file as XLSX.
// Keywords: Aspose.Cells picture reflection | custom reflection gradient C# | set picture opacity Aspose.Cells | reflection size blur distance Aspose.Cells | Excel image reflection .NET | Aspose.Cells ReflectionEffect
// Common Searches: how to add a custom reflection to an image in Aspose.Cells C# | Aspose.Cells set picture reflection size and opacity | configure blur and distance for picture reflection in Excel using .NET | Aspose.Cells reflection direction and fade direction example | C# code for custom picture reflection gradient in Aspose.Cells
// Developer Intent: Add a picture to a worksheet and apply a custom reflection gradient with precise opacity, size, blur, distance, and direction settings using Aspose.Cells for .NET.
// Use Cases: Product catalogs where each product image has a subtle, uniform reflection for visual depth. | Marketing reports that automatically style inserted images with a consistent reflection effect. | Automated presentation‑style Excel sheets where pictures appear to float using the same reflection parameters.
// AI Prompts: Generate C# code that changes the reflection gradient to vertical and sets opacity to 50 % for a picture in Aspose.Cells. | Show how to load reflection settings from a JSON file and apply them to multiple pictures in a workbook with Aspose.Cells for .NET. | Explain how to combine a custom reflection gradient with a shadow effect on a picture shape using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsReflectionDemo
{
    // Creates a workbook, inserts a JPEG picture, and configures Aspose.Cells' ReflectionEffect (Custom) – size, opacity, blur, distance, direction and fade direction – before saving the file as XLSX.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file to be inserted
                string imagePath = "sample.jpg";

                // Verify that the image file exists before adding it
                if (File.Exists(imagePath))
                {
                    // Add the picture to the worksheet (row 2, column 2)
                    int pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
                    Picture picture = worksheet.Pictures[pictureIndex];

                    // Access and configure the reflection effect
                    ReflectionEffect reflection = picture.Reflection;
                    reflection.Type = ReflectionEffectType.Custom;
                    reflection.Size = 85;               // 85% of the shape height
                    reflection.Transparency = 0.25;    // 25% transparent (75% opaque)
                    reflection.Blur = 5;               // slight blur
                    reflection.Distance = 8;           // distance from the shape
                    reflection.Direction = 90;         // gradient direction
                    reflection.FadeDirection = 90;
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Save the workbook with the applied (or skipped) reflection effect
                string outputPath = "PictureWithCustomReflection.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
