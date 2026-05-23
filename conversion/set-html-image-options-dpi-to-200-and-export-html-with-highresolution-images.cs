using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ExportHtmlHighResolution
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and add some sample content
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("High‑DPI HTML Export");

            // Add an image to the worksheet if the file exists
            string imagePath = "example.jpg";
            if (File.Exists(imagePath))
            {
                // Add picture at row index 2 (third row), column index 0 (first column)
                sheet.Pictures.Add(2, 0, imagePath);
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
            }

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Set high‑resolution image options
            ImageOrPrintOptions imgOptions = htmlOptions.ImageOptions;
            imgOptions.HorizontalResolution = 200;
            imgOptions.VerticalResolution = 200;

            // Embed images as Base64 to keep everything in a single HTML file
            htmlOptions.ExportImagesAsBase64 = true;

            // Save the workbook as HTML with the high‑resolution image settings
            string outputPath = "high_res_output.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during export: {ex.Message}");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        ExportHtmlHighResolution.Run();
    }
}