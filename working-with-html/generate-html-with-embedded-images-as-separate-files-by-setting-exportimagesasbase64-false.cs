// Title: Export an Excel workbook to HTML with external image files using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export workbook to HTML with images saved as separate files | How to generate an image folder when saving Excel as HTML with Aspose.Cells | Set ExportImagesAsBase64 false in HtmlSaveOptions for Aspose.Cells .NET | Save only the active worksheet to HTML and keep images as files using Aspose.Cells | C# example for Aspose.Cells HTML export with external PNG images
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 false | Excel to HTML conversion with external images C# | Aspose.Cells generate image folder on HTML export | ExportActiveWorksheetOnly HTML Aspose.Cells | C# save workbook as HTML separate image files

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, adds data and a 1x1 PNG picture, configures HtmlSaveOptions with ExportImagesAsBase64 set to false and ExportActiveWorksheetOnly set to true, then saves the workbook as an HTML file. Aspose.Cells creates a companion folder (e.g., ExportedWorkbook_files) containing the image file, and the console outputs the locations of the HTML file and the generated images folder.
class Program
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and add some sample data.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Widget");
            cells["B2"].PutValue(19.99);
            cells["A3"].PutValue("Gadget");
            cells["B3"].PutValue(29.99);

            // -----------------------------------------------------------------
            // 2. Insert a simple PNG image (1x1 pixel) into the sheet.
            // -----------------------------------------------------------------
            // PNG data for a 1x1 transparent pixel.
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

            using (MemoryStream imgStream = new MemoryStream(pngBytes))
            {
                // Insert picture at cell C1 (row 0, column 2)
                sheet.Pictures.Add(0, 2, imgStream);
            }

            // -----------------------------------------------------------------
            // 3. Configure HTML save options to export images as separate files.
            // -----------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // When false, images are saved as separate files in a folder.
                ExportImagesAsBase64 = false,

                // Export only the active worksheet to keep the example simple.
                ExportActiveWorksheetOnly = true
            };

            // -----------------------------------------------------------------
            // 4. Save the workbook as HTML.
            // -----------------------------------------------------------------
            string outputHtmlPath = "ExportedWorkbook.html";
            workbook.Save(outputHtmlPath, htmlOptions);

            // -----------------------------------------------------------------
            // 5. Inform the user where the files were created.
            // -----------------------------------------------------------------
            Console.WriteLine($"HTML file saved to: {Path.GetFullPath(outputHtmlPath)}");

            // The images are stored in a folder that Aspose creates next to the HTML file.
            // Its name is derived from the HTML file name (e.g., ExportedWorkbook_files).
            string imagesFolder = Path.Combine(
                Path.GetDirectoryName(Path.GetFullPath(outputHtmlPath)) ?? "",
                Path.GetFileNameWithoutExtension(outputHtmlPath) + "_files");

            if (Directory.Exists(imagesFolder))
            {
                Console.WriteLine($"Images folder created at: {imagesFolder}");
                Console.WriteLine("Contained image files:");
                foreach (string file in Directory.GetFiles(imagesFolder))
                {
                    Console.WriteLine("  " + Path.GetFileName(file));
                }
            }
            else
            {
                Console.WriteLine("No images folder was created. Check if ExportImagesAsBase64 was set correctly.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during the HTML export process:");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}
