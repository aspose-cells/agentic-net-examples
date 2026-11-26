// Title: Apply a custom reflection gradient with opacity, size, blur, and direction to a picture shape using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set custom reflection gradient on a worksheet picture | how to adjust reflection opacity and size for a picture shape in Aspose.Cells | configure reflection blur and distance for an Excel image using Aspose.Cells | C# example of applying custom reflection effect with direction and fade to a picture | set reflection direction and fade direction for a picture shape in Aspose.Cells workbook
// Tags: reflection size transparency Aspose.Cells | reflection blur distance Aspose.Cells | reflection direction fade Aspose.Cells | picture shape reflection customization Aspose.Cells | Excel image reflection C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a temporary 1x1 red PNG, inserts it as a picture into the first worksheet of a new workbook, and configures the picture's ReflectionEffect to Custom with specific size, transparency, blur, distance, direction, fade direction, and rotation settings before saving the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        // Path for the temporary image file (a 1x1 red PNG)
        string tempImagePath = Path.Combine(Path.GetTempPath(), "tempImage.png");

        // Base64 representation of a 1x1 red pixel PNG
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";

        try
        {
            // Write the image to disk
            File.WriteAllBytes(tempImagePath, Convert.FromBase64String(base64Png));

            // Verify that the image file was created
            if (!File.Exists(tempImagePath))
                throw new FileNotFoundException("Temporary image file was not created.", tempImagePath);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add the picture to the worksheet using the file path overload
            int pictureIndex = worksheet.Pictures.Add(1, 1, tempImagePath);
            Picture pictureShape = worksheet.Pictures[pictureIndex];

            // Apply a custom reflection gradient
            ReflectionEffect reflection = pictureShape.Reflection;
            reflection.Type = ReflectionEffectType.Custom;   // Enable custom reflection
            reflection.Size = 80;                           // End position of the alpha ramp (percentage)
            reflection.Transparency = 0.3;                  // Starting transparency (0 = opaque, 1 = fully transparent)
            reflection.Blur = 5;                            // Blur radius in points
            reflection.Distance = 10;                       // Distance from the shape in points
            reflection.Direction = 90;                      // Direction of the alpha gradient ramp (degrees)
            reflection.FadeDirection = 90;                  // Direction to offset the reflection (degrees)
            reflection.RotWithShape = true;                 // Make reflection rotate with the shape

            // Save the workbook
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomReflectionGradient.xlsx");
            workbook.Save(outputPath);

            Console.WriteLine("Workbook saved successfully to:");
            Console.WriteLine(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
        finally
        {
            // Clean up the temporary image file
            if (File.Exists(tempImagePath))
            {
                try { File.Delete(tempImagePath); } catch { /* ignore cleanup errors */ }
            }
        }
    }
}
