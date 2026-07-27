// Title: Apply a Default Shadow to Every Inserted Picture in Aspose.Cells (C#)
// Description: Demonstrates how to create a workbook, insert an image, and automatically apply a predefined ShadowEffect (preset type, transparency, size, distance) to each new picture before saving the file.
// Keywords: Aspose.Cells | C# | picture shadow | ShadowEffect | preset shadow type | Excel image styling | automatic picture formatting | workbook picture insertion | .NET Excel API
// Common Searches: Aspose.Cells add picture with shadow | C# set picture shadow effect Aspose.Cells | default shadow for inserted images Aspose.Cells | apply preset shadow to pictures in Excel using Aspose | automate picture styling Aspose.Cells .NET
// Developer Intent: Automatically apply a predefined shadow style to every picture shape when it is added to a worksheet.
// Use Cases: Add product photos to sales reports with a consistent shadow for visual depth. | Insert company logos into marketing brochures where each logo receives the same shadow styling. | Generate dashboards that display chart screenshots with a uniform shadow effect across all images.
// AI Prompts: Write a C# method that applies a customizable ShadowEffect to an Aspose.Cells Picture, exposing parameters for preset type, transparency, size, and distance. | Show how to wrap the Pictures.Add call so that the default shadow is automatically applied to each newly inserted picture. | Provide code to iterate over all pictures in a worksheet and apply the same predefined shadow settings.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowDemo
{
    // Demonstrates how to create a workbook, insert an image, and automatically apply a predefined ShadowEffect (preset type, transparency, size, distance) to each new picture before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Path to the image file to be inserted
                string imagePath = "example.jpg";

                // Verify that the image file exists before adding it
                if (File.Exists(imagePath))
                {
                    // Add a picture to the worksheet (row 2, column 2)
                    int pictureIndex = sheet.Pictures.Add(2, 2, imagePath);
                    Picture picture = sheet.Pictures[pictureIndex];

                    // Apply a predefined shadow style to the newly added picture
                    ApplyDefaultShadow(picture);
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
                }

                // Save the workbook
                string outputPath = "PictureWithShadow.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <param name="picture">The picture shape to style.</param>
        private static void ApplyDefaultShadow(Picture picture)
        {
            // Access the ShadowEffect object of the shape
            ShadowEffect shadow = picture.ShadowEffect;

            // Set a preset shadow type (e.g., OffsetDiagonalBottomRight)
            shadow.PresetType = PresetShadowType.OffsetDiagonalBottomRight;

            // Customize additional shadow properties
            shadow.Transparency = 0.3f; // 30% transparent
            shadow.Size = 100;          // Size as a percentage
            shadow.Distance = 5;        // Distance from the shape
            // Note: ShadowEffect does not have a Direction property; it is defined by the preset type.
        }
    }
}
