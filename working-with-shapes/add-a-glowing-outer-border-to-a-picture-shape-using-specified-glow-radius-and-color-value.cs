// Title: Add an outer glow border to a picture shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add outer glow to picture shape in Excel | apply custom glow effect to image before inserting with Aspose.Cells | set picture border color and thickness in Aspose.Cells worksheet | workaround for missing EffectFormat on picture shapes in Aspose.Cells .NET | preprocess PNG with glow then insert into Excel using Aspose.Cells
// Tags: Aspose.Cells picture outer glow | C# add picture border in Excel | preprocess image for glow before Aspose.Cells insertion | Aspose.Cells effect format limitation | Excel image glow workaround .NET

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, inserts a PNG as a picture shape on the first worksheet, explains that Aspose.Cells for .NET does not expose an EffectFormat for pictures, and shows where custom image‑processing code could be added to apply an outer glow with a specific radius and color before the picture is added to the worksheet. The workbook is then saved as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the image file to be inserted as a picture shape
            string imagePath = "sample.png";

            // Verify that the image file exists before adding
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Add the picture to the worksheet (positioned at row 0, column 0)
            int pictureIndex = worksheet.Pictures.Add(0, 0, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // NOTE: Aspose.Cells for .NET does not expose an EffectFormat for Picture/Shape objects.
            // If a glow effect is required, it must be applied via external image processing before insertion.
            // The following placeholder demonstrates where such logic would be placed.
            // ApplyGlowEffect(picture, radius: 10f, color: Color.FromArgb(255, 255, 0, 0));

            // Save the workbook
            string outputPath = "output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Placeholder for a custom glow effect implementation (not part of Aspose.Cells API)
    // private static void ApplyGlowEffect(Picture picture, float radius, Color color)
    // {
    //     // Implement image processing here if needed.
    // }
}
