using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Path to the image file
                string imagePath = "example.jpg";

                // Add image if the file exists
                if (File.Exists(imagePath))
                {
                    sheet.Pictures.Add(0, 0, imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping image insertion.");
                }

                // Configure HTML save options to embed images as Base64
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true
                };

                // Save the workbook as an HTML file
                string outputPath = "output.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"HTML file '{outputPath}' saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}