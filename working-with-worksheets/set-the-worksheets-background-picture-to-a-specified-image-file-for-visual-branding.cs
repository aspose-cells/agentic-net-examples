// Title: Add a PNG image as a worksheet background in Aspose.Cells for .NET (C#)
// Common Searches: how to set a worksheet background image using Aspose.Cells C# | Aspose.Cells add picture to entire sheet programmatically | C# insert branding logo as Excel worksheet background with Aspose | using Aspose.Cells picture covering sheet as background example | load external PNG as background for Excel worksheet in .NET
// Tags: worksheet.Pictures.Add background image | Aspose.Cells set worksheet background PNG | add sheet‑wide picture Aspose.Cells C# | brand Excel workbook with image using Aspose | load external image as worksheet background .NET

using Aspose.Cells;
using System;
using System.IO;

// The example verifies a PNG file, creates a new Workbook, accesses the first Worksheet, inserts the image at cell A1 with Worksheet.Pictures.Add to act as a sheet‑wide background, and saves the workbook as BrandedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the image that will be used as the worksheet background
            string imagePath = @"C:\Images\brand.png";

            // Verify that the image file exists before attempting to load it
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add the image as a picture covering the sheet (acts as a background)
            // Position at cell A1 (row 0, column 0)
            worksheet.Pictures.Add(0, 0, imagePath);

            // Save the workbook with the applied picture
            string outputPath = "BrandedWorkbook.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
