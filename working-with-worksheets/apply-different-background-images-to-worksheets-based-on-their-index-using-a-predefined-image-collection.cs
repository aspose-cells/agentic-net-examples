// Title: Apply distinct background images to each worksheet in an Aspose.Cells workbook using C#
// Common Searches: C# Aspose.Cells set a different background picture for each worksheet | How to add rotating background images to Excel sheets with Aspose.Cells | Assign background images to worksheets based on sheet index in C# | Using Pictures.Add to insert worksheet background images in Aspose.Cells | Load multiple background PNG files and apply them to Excel worksheets with Aspose.Cells | Save an Aspose.Cells workbook with custom sheet background pictures
// Tags: worksheet background picture Aspose.Cells | add picture to worksheet C# | rotate sheet background images Aspose | load background images from file list C# | save workbook with custom sheet backgrounds Aspose.Cells | picture anchored at cell A1 Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, adds several worksheets, and defines a list of image file paths. It loops through each worksheet, selects an image using the sheet index modulo the image count, checks that the file exists, and inserts the image as a picture anchored at cell A1 via sheet.Pictures.Add. Finally, the workbook is saved as an XLSX file to the specified output directory.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add worksheets for demonstration
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");
            workbook.Worksheets.Add("Sheet4");

            // Collection of background image file paths
            List<string> backgroundImages = new List<string>
            {
                @"C:\Images\Background1.png",
                @"C:\Images\Background2.png",
                @"C:\Images\Background3.png"
                // Add more image paths as needed
            };

            // Verify that at least one image exists
            if (backgroundImages.Count == 0)
                throw new FileNotFoundException("No background image paths were provided.");

            // Apply an image to each worksheet (as a picture to avoid System.Drawing dependency)
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                string imagePath = backgroundImages[i % backgroundImages.Count];

                // Ensure the image file exists before loading
                if (!File.Exists(imagePath))
                    throw new FileNotFoundException($"Background image not found: {imagePath}");

                try
                {
                    // Add the image as a picture anchored at the top‑left corner (cell A1)
                    sheet.Pictures.Add(0, 0, imagePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add image to sheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Prepare output directory
            string outputPath = @"C:\Output\WorkbookWithBackgrounds.xlsx";
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
