using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class EnableCssCustomPropertiesDemo
    {
        public static void Main()
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add header text with bold blue font
            sheet.Cells["A1"].PutValue("Header");
            Style headerStyle = sheet.Cells["A1"].GetStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.Blue;
            sheet.Cells["A1"].SetStyle(headerStyle);

            // Add some sample data
            sheet.Cells["A2"].PutValue("Data 1");
            sheet.Cells["B2"].PutValue("Data 2");

            // Path to the image file
            string imagePath = "logo.png";

            // Insert an image that will appear multiple times (if the file exists)
            if (File.Exists(imagePath))
            {
                int imgIdx = sheet.Pictures.Add(2, 0, imagePath);
                Picture pic = sheet.Pictures[imgIdx];
                pic.Width = 80;
                pic.Height = 80;

                // Copy the same image to another location
                int imgIdx2 = sheet.Pictures.Add(5, 2, imagePath);
                Picture pic2 = sheet.Pictures[imgIdx2];
                pic2.Width = 80;
                pic2.Height = 80;
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                EnableCssCustomProperties = true, // Optimize CSS using custom properties
                TableCssId = "custom-table"       // Prefix for table CSS classes
            };

            // Save the workbook as HTML with the configured options
            string outputPath = "OutputWithCustomProperties.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}