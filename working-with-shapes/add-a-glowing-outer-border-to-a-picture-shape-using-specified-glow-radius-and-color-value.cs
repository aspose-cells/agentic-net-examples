// Title: Add an orange outer glow to a picture shape in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Insert an image as a picture shape into a worksheet and apply an orange glow with a radius of 8 using Aspose.Cells C#. | Use reflection to access the EffectFormat of a picture shape and set its Glow radius and color when the property is not directly exposed. | Save the workbook after modifying the picture's glow effect to generate the final Excel file.
// Common Searches: how to add outer glow to an inserted picture in Aspose.Cells C# | Aspose.Cells set picture glow radius and color using reflection | apply orange glow effect to Excel image with Aspose.Cells .NET example
// Tags: picture shape glow Aspose.Cells C# | set glow radius color EffectFormat | reflection workaround EffectFormat Aspose.Cells | outer glow border Excel image | apply orange glow to worksheet picture

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an image, inserts it as a picture shape into the first worksheet, uses reflection to reach the picture's EffectFormat and sets the Glow radius to 8 with an orange color, then saves the workbook as GlowingPicture.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the image file to be inserted
                string imagePath = "image.png";

                // Verify that the image file exists to avoid FileNotFoundException
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {Path.GetFullPath(imagePath)}");
                    return;
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a picture shape to the worksheet at row 2, column 2 (zero‑based indices)
                int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);

                // Retrieve the picture object
                Picture picture = worksheet.Pictures[pictureIndex];

                // Attempt to apply a glow effect using reflection (EffectFormat may not be available in older versions)
                try
                {
                    var effectProp = picture.GetType().GetProperty("EffectFormat");
                    if (effectProp != null)
                    {
                        var effectFormat = effectProp.GetValue(picture);
                        var glowProp = effectFormat?.GetType().GetProperty("Glow");
                        if (glowProp != null)
                        {
                            var glow = glowProp.GetValue(effectFormat);
                            var radiusProp = glow?.GetType().GetProperty("Radius");
                            var colorProp = glow?.GetType().GetProperty("Color");
                            if (radiusProp != null && colorProp != null)
                            {
                                radiusProp.SetValue(glow, 8); // example radius
                                colorProp.SetValue(glow, Color.FromArgb(255, 255, 165, 0)); // orange color
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Glow effect could not be applied: {ex.Message}");
                }

                // Save the workbook to a file
                string outputPath = "GlowingPicture.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
