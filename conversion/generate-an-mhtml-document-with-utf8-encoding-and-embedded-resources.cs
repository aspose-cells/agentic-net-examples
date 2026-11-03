// Title: Create an MHTML file with UTF-8 encoding and embedded images from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to export an Excel workbook to MHTML with UTF-8 encoding using Aspose.Cells in C# | Aspose.Cells C# save worksheet as .mhtml with embedded base64 images | generate MHTML document from Excel and include PNG picture using Aspose.Cells HtmlSaveOptions | C# Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 example for MHTML output | save Excel data as MHTML file with embedded resources using Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions MHTML export | C# export Excel to MHTML with UTF-8 | embed images as base64 in MHTML using Aspose.Cells | save workbook as .mhtml Aspose.Cells .NET | UTF-8 encoding for MHTML output Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

// The example creates a workbook, fills cells with text, date and numeric values, optionally adds a PNG picture, configures HtmlSaveOptions for UTF-8 encoding and Base64 image embedding, and saves the result as an MHTML document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sample";

            // Populate some data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B2"].PutValue(12345);

            // Insert an image if the file exists
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logo.png");
            if (File.Exists(imagePath))
            {
                int pictureIndex = sheet.Pictures.Add(2, 0, imagePath);
                sheet.Pictures[pictureIndex].Placement = PlacementType.FreeFloating;
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping image insertion.");
            }

            // Configure HTML save options (used for MHTML output) with UTF-8 encoding and embedded images
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                Encoding = Encoding.UTF8,
                ExportImagesAsBase64 = true // embed images directly in the HTML/MHTML
            };

            // Determine output path and ensure the directory exists
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sample.mhtml");
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an MHTML document (HTML content with .mhtml extension)
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
