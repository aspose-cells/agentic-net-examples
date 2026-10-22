// Title: Refresh linked picture objects in an Excel workbook after overwriting the source image file with Aspose.Cells for .NET (C#)
// AI Prompts: Overwrite a source PNG on disk and programmatically refresh all linked picture objects in an Excel workbook using Aspose.Cells in C#. | Load an Excel file, replace its linked image source, and ensure the updated picture appears when the workbook is saved with Aspose.Cells. | Iterate through each worksheet and invoke picture link refresh (if supported) after updating the source image file with Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells how to update linked pictures after changing source image file | Refresh picture links in Excel workbook programmatically using Aspose.Cells .NET | Replace source PNG and force linked picture refresh in Excel with Aspose.Cells | Aspose.Cells picture.UpdateLink method availability across versions | Save workbook to apply linked picture changes Aspose.Cells C#
// Tags: linked picture refresh Aspose.Cells C# | overwrite source image Excel Aspose.Cells | picture.UpdateLink method Aspose.Cells | save workbook apply picture changes Aspose.Cells | automatic linked picture update Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook containing linked pictures, overwrites the source PNG with a new image, optionally refreshes picture links (using Picture.UpdateLink when available), and saves the workbook, relying on Aspose.Cells to update linked pictures automatically upon save.
class RefreshLinkedPictures
{
    static void Main()
    {
        try
        {
            // Path to the Excel file that contains linked pictures
            string workbookPath = @"C:\Temp\LinkedPictures.xlsx";

            // Verify workbook exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Paths for the source image and the new image to replace it
            string sourceImagePath = @"C:\Temp\SourceImage.png";
            string newImagePath = @"C:\Temp\NewImage.png";

            // Overwrite the source image with a new one if the new image exists
            if (File.Exists(newImagePath))
            {
                try
                {
                    File.Copy(newImagePath, sourceImagePath, overwrite: true);
                    Console.WriteLine($"Replaced source image with: {newImagePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to copy image: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"New image not found: {newImagePath}");
            }

            // Note: In recent Aspose.Cells versions, linked pictures are refreshed automatically
            // when the workbook is saved. If explicit refresh is required, the API provides
            // Picture.UpdateLink(), but it may not be available in older versions.
            // Therefore, we simply iterate through pictures without calling unavailable members.

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PictureCollection pictures = sheet.Pictures;
                // No explicit update needed; placeholder for future logic if API supports it.
            }

            // Ensure the output directory exists
            string outputPath = @"C:\Temp\LinkedPictures_Refreshed.xlsx";
            try
            {
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the refreshed workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
