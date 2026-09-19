// Title: Create a picture shape with a custom reflection gradient (size, opacity, blur, distance) in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a PNG file as a picture shape into a worksheet and set Reflection.Size to 0.6, Reflection.Transparency to 0.3, Reflection.Blur to 0.2, and Reflection.Distance to 0.05 with Aspose.Cells in C#. | Generate an XLSX workbook that demonstrates a gradient reflection effect on an inserted image by adjusting the Shape.Reflection properties for size, opacity, blur, and distance. | Write C# code that adds an image to a worksheet, applies a custom reflection gradient, and saves the workbook as a .xlsx file using Aspose.Cells.
// Common Searches: Aspose.Cells C# set picture shape reflection size and transparency | How to add gradient reflection to an image in Excel using Aspose.Cells | C# Aspose.Cells Reflection.Blur and Reflection.Distance example | Create custom reflection effect for picture shape in Aspose.Cells workbook | Adjust reflection opacity and size of inserted picture in Aspose.Cells
// Tags: Aspose.Cells picture shape reflection gradient | set reflection size transparency Aspose.Cells | configure reflection blur distance C# | add image with gradient reflection Excel | Shape.Reflection properties Aspose.Cells | custom reflection effect workbook C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a PNG image as a picture shape, configures the shape's Reflection properties (Size, Transparency, Blur, Distance) to produce a gradient reflection effect, and saves the file as ReflectionGradientDemo.xlsx.
class ApplyReflectionGradient
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define picture placement coordinates
            int upperLeftRow = 2;
            int upperLeftColumn = 2;
            int lowerRightRow = 12;
            int lowerRightColumn = 12;

            // Path to the picture file (ensure the file exists)
            string picturePath = "sample-image.png";

            if (File.Exists(picturePath))
            {
                // Load the picture into a stream (required by AddPicture overload)
                using (FileStream pictureStream = new FileStream(picturePath, FileMode.Open, FileAccess.Read))
                {
                    // Add the picture shape to the worksheet
                    Shape pictureShape = sheet.Shapes.AddPicture(
                        upperLeftRow, upperLeftColumn,
                        lowerRightRow, lowerRightColumn,
                        pictureStream);

                    // Configure the reflection gradient
                    // Note: In current Aspose.Cells versions the reflection type defaults to Gradient,
                    // so setting the Type property is not required and the enum may be unavailable.
                    pictureShape.Reflection.Size = 0.6;          // 60% of the original shape height
                    pictureShape.Reflection.Transparency = 0.3; // 30% transparent
                    pictureShape.Reflection.Blur = 0.2;         // softness of the reflection
                    pictureShape.Reflection.Distance = 0.05;   // gap between shape and its reflection
                }
            }
            else
            {
                Console.WriteLine($"Picture file not found: {picturePath}. Skipping picture insertion.");
            }

            // Save the workbook to a file
            try
            {
                workbook.Save("ReflectionGradientDemo.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
