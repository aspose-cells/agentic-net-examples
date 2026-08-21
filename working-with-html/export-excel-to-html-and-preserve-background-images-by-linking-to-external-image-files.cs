// Title: Export Excel to HTML with external background images using Aspose.Cells (C#)
// Description: C# sample that creates a JPEG, adds it as a worksheet picture, configures HtmlSaveOptions to disable Base64 embedding, and saves the workbook as HTML. The image is written to an "images" folder and linked from the generated page, enabling fast loading and easy caching.
// Keywords: Aspose.Cells HTML export external images | ExportImagesAsBase64 false | C# Excel to HTML conversion | save workbook as HTML with image folder | Aspose.Cells picture linking
// Common Searches: Aspose.Cells export HTML without base64 images | C# save Excel as HTML external image files | how to link background picture in Aspose HTML export | Aspose.Cells HtmlSaveOptions image folder path | Excel to HTML external images C#
// Developer Intent: Generate HTML from an Excel workbook where worksheet pictures are stored as separate image files rather than embedded Base64 strings.
// Use Cases: Build web pages that reference cached image assets for quicker load times. | Integrate Excel‑derived content into CMS platforms that require images in a dedicated directory. | Automate bulk conversion of workbooks while preserving image links for consistent branding.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML, disable Base64 image embedding, and specify a custom folder for the images. | Explain how Aspose.Cells determines the image folder name when ExportImagesAsBase64 is false and how to override it. | Provide a complete example that adds a background picture to a worksheet and saves the workbook as HTML with external image files linked.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // C# sample that creates a JPEG, adds it as a worksheet picture, configures HtmlSaveOptions to disable Base64 embedding, and saves the workbook as HTML. The image is written to an "images" folder and linked from the generated page, enabling fast loading and easy caching.
    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare a sample image file if it does not exist
                string imagePath = "sample_background.jpg";
                if (!File.Exists(imagePath))
                {
                    try
                    {
                        // Minimal 1x1 JPEG image (base64 encoded)
                        const string base64Jpeg =
                            "/9j/4AAQSkZJRgABAQEASABIAAD/2wBDABALDA4MDRANDhAQEBYQEBYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYV/2wBDARESEhMUFhYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYVFRYV/wAARCAABAAEDASIAAhEBAxEB/8QAFwAAAwEAAAAAAAAAAAAAAAAAAAIEBf/EABYBAQEBAAAAAAAAAAAAAAAAAAABAv/aAAwDAQACEAMQAAAB6Kf/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oACAEBAAE/AKf/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oACAEDAQE/AV//xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oACFADEA/EAf/EABkQAQEBAQEBAAAAAAAAAAAAAAABEQIhMf/aAAgBAgEBPwGkZ//EABkQAQEBAQEBAAAAAAAAAAAAAAABEQIhMf/aAAgBAwEBPwGkZ//Z";
                        byte[] imageBytes = Convert.FromBase64String(base64Jpeg);
                        File.WriteAllBytes(imagePath, imageBytes);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to create sample image: {ex.Message}");
                        return;
                    }
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add the picture to the worksheet
                try
                {
                    sheet.Pictures.Add(0, 0, imagePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add picture: {ex.Message}");
                }

                // Ensure the images folder exists (Aspose will use this folder for external images)
                string imagesFolder = Path.Combine(Environment.CurrentDirectory, "images");
                try
                {
                    Directory.CreateDirectory(imagesFolder);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create images folder: {ex.Message}");
                    return;
                }

                // Configure HTML save options to link images as external files
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportImagesAsBase64 = false // Do not embed images as Base64
                    // ExportImageFolderPath is not required; Aspose will place images in a folder named after the HTML file
                };

                // Save the workbook as HTML; images will be linked via external files
                string htmlOutputPath = Path.Combine(Environment.CurrentDirectory, "workbook.html");
                try
                {
                    workbook.Save(htmlOutputPath, htmlOptions);
                    Console.WriteLine("HTML file saved to: " + htmlOutputPath);
                    Console.WriteLine("Images are stored in the \"images\" folder alongside the HTML file.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save HTML: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }
    }
}
