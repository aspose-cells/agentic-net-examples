using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HtmlExportBase64Demo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data
            worksheet.Cells["A1"].PutValue("Sample Image Below:");

            // Path to the image file
            string imagePath = "example.jpg";

            // Add an image if the file exists
            if (File.Exists(imagePath))
            {
                worksheet.Pictures.Add(2, 0, imagePath);
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
            }

            // Configure HTML save options to embed images as Base64
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true,   // Embed images directly in the <img> tag
                SaveAsSingleFile = true        // Produce a single HTML file with all resources embedded
            };

            // Save the workbook as a single HTML file with embedded images
            string outputPath = "output.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"HTML file saved successfully at: {outputPath}");
        }
    }
}