// Title: Apply a custom reflection effect with opacity and size to a picture shape using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom reflection opacity for an inserted picture | How to change the reflection size of a picture shape in an Excel workbook with Aspose.Cells | C# code to apply a reflection effect with transparency to a worksheet image using Aspose.Cells | Adjust picture shape reflection distance and blur in Aspose.Cells for .NET
// Tags: picture shape reflection opacity Aspose.Cells | custom reflection size Excel image C# | set picture reflection properties Aspose.Cells | adjust reflection distance blur worksheet picture | Aspose.Cells reflection effect on picture shape

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a picture into a worksheet, accesses its ReflectionEffect, sets the reflection's Transparency to 0.3 and Size to 0.5 (with optional Distance and Blur adjustments), and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the picture file exists to avoid FileNotFoundException
            string picturePath = "image.png";
            if (!File.Exists(picturePath))
            {
                Console.WriteLine($"Picture file not found: {picturePath}");
                return;
            }

            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a picture shape to the worksheet (upper left row/column, lower right row/column)
            // Adjust the cell range as needed for the picture size
            int pictureIndex = sheet.Pictures.Add(2, 2, 10, 10, picturePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Apply a custom reflection effect to the picture
            ReflectionEffect reflection = picture.Reflection;

            // Set the reflection type to Custom if the enum is available.
            // If the enum is not present in the referenced version, the default type is already Custom.
            // reflection.Type = Aspose.Cells.Drawing.ReflectionType.Custom;

            // Adjust opacity (Transparency) – value between 0 (opaque) and 1 (fully transparent)
            reflection.Transparency = 0.3; // 30% transparent

            // Adjust size – value between 0 (no reflection) and 1 (full size of the original shape)
            reflection.Size = 0.5; // 50% of the original shape's height

            // Optional: fine‑tune other reflection properties if desired
            reflection.Distance = 0.0; // distance from the original shape
            reflection.Blur = 0.0;     // blur amount

            // Save the workbook (lifecycle rule: save)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully as {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
