using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsImageQualityValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(150);

            // Configure image rendering options for JPEG with a specific quality
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Jpeg,
                Quality = 80 // Quality range: 0-100
            };

            // Render the worksheet to a JPEG image
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            string imagePath = "ExportedSheet.jpg";
            renderer.ToImage(0, imagePath);

            // Validate the exported JPEG file size
            const long minSizeBytes = 50 * 1024;   // 50 KB
            const long maxSizeBytes = 500 * 1024;  // 500 KB

            FileInfo fileInfo = new FileInfo(imagePath);
            long fileSize = fileInfo.Length;

            Console.WriteLine($"Exported JPEG size: {fileSize} bytes");

            if (fileSize < minSizeBytes)
            {
                Console.WriteLine("Image size is too small; quality may be insufficient.");
            }
            else if (fileSize > maxSizeBytes)
            {
                Console.WriteLine("Image size is too large; quality may be higher than needed.");
            }
            else
            {
                Console.WriteLine("Image size is within the acceptable range.");
            }
        }
    }
}