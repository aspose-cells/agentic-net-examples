// Title: Apply Custom Reflection Gradient to a Picture Shape with Aspose.Cells for .NET
// Description: Shows how to add an image to an Excel worksheet using Aspose.Cells, retrieve its Shape, set the ReflectionEffect to Custom, and adjust opacity, size, blur, distance, direction and other reflection parameters before saving the workbook.
// Keywords: Aspose.Cells | C# | .NET | ReflectionEffect | custom reflection | picture shape | image opacity | reflection transparency | reflection size | gradient reflection | Excel workbook | add picture to worksheet | shape formatting | reflection blur | reflection distance | reflection direction
// Common Searches: Aspose.Cells set custom reflection opacity | C# picture shape reflection size Aspose.Cells | how to adjust reflection blur distance direction in Excel using Aspose | apply gradient reflection to image shape Aspose.Cells .NET | save workbook after modifying picture reflection
// Developer Intent: Add a picture to an Excel sheet and apply a custom reflection effect, controlling its transparency, size and other visual attributes via Aspose.Cells.
// Use Cases: Design marketing dashboards where product photos have a subtle reflective finish. | Generate product catalogs with images that mimic a glossy surface using custom reflections. | Create slide‑style Excel presentations where each picture includes a tailored reflection to match branding guidelines.
// AI Prompts: Write C# code with Aspose.Cells that inserts a picture and sets a custom reflection with 40% transparency and 70% size, then saves the file. | Show how to modify reflection blur, distance, and direction for a picture shape in Aspose.Cells for .NET. | Explain how to read and update the ReflectionEffect properties of an existing picture shape in a saved workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsReflectionDemo
{
    // Shows how to add an image to an Excel worksheet using Aspose.Cells, retrieve its Shape, set the ReflectionEffect to Custom, and adjust opacity, size, blur, distance, direction and other reflection parameters before saving the workbook.
    public class ApplyCustomReflectionToPicture
    {
        public static void Run()
        {
            try
            {
                // Verify that the source image exists
                const string imagePath = "sample.jpg";
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Operation aborted.");
                    return;
                }

                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a picture to the worksheet
                // Parameters: upper left row, upper left column, picture file name
                int pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Access the shape that represents the picture
                Shape pictureShape = picture;

                // Obtain the ReflectionEffect object for the picture shape
                ReflectionEffect reflection = pictureShape.Reflection;

                // Set the reflection type to Custom to allow manual adjustment of properties
                reflection.Type = ReflectionEffectType.Custom;

                // Adjust opacity (Transparency) – 0.0 = fully opaque, 1.0 = fully transparent
                reflection.Transparency = 0.3; // 30% transparent (70% opaque)

                // Adjust the size of the reflection – value is a percentage (0‑100)
                reflection.Size = 80; // 80% of the original shape height

                // Optional: fine‑tune other visual aspects
                reflection.Blur = 5;          // Softens the reflection edges
                reflection.Distance = 10;     // Moves the reflection away from the shape
                reflection.Direction = 90;    // Gradient direction (degrees)
                reflection.FadeDirection = 90;
                reflection.RotWithShape = true;

                // Save the workbook (lifecycle: save)
                const string outputPath = "PictureWithCustomReflection.xlsx";
                workbook.Save(outputPath);

                // Output confirmation
                Console.WriteLine($"Custom reflection applied and workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyCustomReflectionToPicture.Run();
        }
    }
}
