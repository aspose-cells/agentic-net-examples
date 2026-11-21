// Title: Add a linked picture from an external URL to an Excel worksheet using Aspose.Cells for .NET without embedding image data
// Common Searches: Aspose.Cells C# add picture from web URL without embedding | How to create a linked image in Excel using Aspose.Cells .NET | Save Excel file with external image reference using Aspose.Cells | Set picture.IsLinked true in Aspose.Cells C# example | Insert external PNG as linked picture in workbook with Aspose.Cells
// Tags: linked picture insertion Aspose.Cells | picture.IsLinked usage .NET | external image reference in Excel workbook | save workbook without embedding images | insert picture from external URL

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating a new Workbook, adding a picture linked to an external URL via sheet.Pictures.Add, optionally naming the picture, ensuring the output directory exists, and saving the workbook as XLSX without embedding the image data.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // URL of the image to link
            string imageUrl = "https://example.com/image.png";

            // Add a linked picture directly from the URL
            int pictureIndex = sheet.Pictures.Add(0, 0, imageUrl);
            Picture picture = sheet.Pictures[pictureIndex];
            // Optional: set a name for the picture
            picture.Name = "LinkedImage";

            // Determine output path and ensure its directory exists
            string outputPath = "LinkedPicture.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
