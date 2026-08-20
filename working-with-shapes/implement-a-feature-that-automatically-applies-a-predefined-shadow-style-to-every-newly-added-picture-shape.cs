// Title: Apply a Default Shadow Style to Every Inserted Picture in Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, inserts a picture, automatically sets a preset diagonal shadow with 30% transparency and 100% size, and saves the file.
// Keywords: Aspose.Cells picture shadow C# | auto apply shadow Aspose.Cells | preset shadow type worksheet image | default picture shadow .NET | Aspose.Cells shadow effect sample | C# insert image with shadow | Aspose.Cells GitHub example | worksheet picture formatting
// Common Searches: how to add a shadow to pictures in Aspose.Cells | Aspose.Cells preset shadow effect C# example | set default picture shadow in .NET workbook | auto apply shadow to inserted images Aspose.Cells | Aspose.Cells picture shadow transparency size
// Developer Intent: Automatically apply a predefined shadow style to each picture added to a worksheet.
// Use Cases: Add a company logo to multiple reports and ensure every instance has the same diagonal shadow. | Generate a product catalog where each product image receives a uniform shadow for visual consistency. | Create a template that inserts user‑uploaded images with a default shadow, removing the need for manual formatting.
// AI Prompts: Write C# code using Aspose.Cells that inserts a picture and applies a preset diagonal shadow with custom transparency and size. | Provide a reusable method that intercepts picture insertion in a workbook and sets a default shadow style for all new pictures.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, inserts a picture, automatically sets a preset diagonal shadow with 30% transparency and 100% size, and saves the file.
class ApplyShadowToPictures
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the image file to be inserted as a picture
            string imagePath = "example.jpg";

            // Verify that the image file exists before adding it
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {Path.GetFullPath(imagePath)}");
                // Optionally, you could exit or continue without adding the picture
                return;
            }

            // Add a picture to the worksheet (row 2, column 2)
            int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Apply a predefined shadow style to the newly added picture
            picture.ShadowEffect.PresetType = PresetShadowType.OffsetDiagonalBottomRight;
            // Optional additional shadow settings
            picture.ShadowEffect.Transparency = 0.3f;   // 30% transparent
            picture.ShadowEffect.Size = 100;           // Shadow size (percentage)

            // Save the workbook with the picture that now has a shadow effect
            string outputPath = "OutputWithShadow.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
