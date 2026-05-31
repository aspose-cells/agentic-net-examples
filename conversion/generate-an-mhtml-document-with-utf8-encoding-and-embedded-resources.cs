using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsMhtmlExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Aspose.Cells MHTML Export");
            sheet.Cells["B2"].PutValue(DateTime.Now);

            // Optionally add an image to demonstrate embedded resources
            // Ensure the image file exists at the specified path
            string imagePath = "example.jpg";
            if (File.Exists(imagePath))
            {
                sheet.Pictures.Add(2, 2, imagePath);
            }

            // Create HTML save options for MHTML format
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);

            // Set UTF-8 encoding
            saveOptions.Encoding = Encoding.UTF8;

            // Embed images as Base64 strings within the MHTML file
            saveOptions.ExportImagesAsBase64 = true;

            // Save the workbook as an MHTML document with embedded resources
            string outputPath = "output.mht";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"MHTML file saved to: {Path.GetFullPath(outputPath)}");
        }
    }
}