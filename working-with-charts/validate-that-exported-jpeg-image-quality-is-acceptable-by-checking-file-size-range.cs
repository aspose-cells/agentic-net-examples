using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsImageQualityCheck
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

            // Set up image rendering options for JPEG with a specific quality
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Jpeg, // Export as JPEG
                Quality = 80               // Desired JPEG quality (0-100)
            };

            // Render the first worksheet page to a JPEG file
            string imagePath = "WorksheetExport.jpg";
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, imagePath);

            // Validate the exported JPEG file size
            const long minSizeBytes = 50 * 1024;   // 50 KB
            const long maxSizeBytes = 500 * 1024;  // 500 KB

            FileInfo fileInfo = new FileInfo(imagePath);
            long fileSize = fileInfo.Length;

            Console.WriteLine($"Exported JPEG size: {fileSize} bytes");

            if (fileSize < minSizeBytes)
            {
                Console.WriteLine("Warning: JPEG file size is smaller than the acceptable minimum.");
            }
            else if (fileSize > maxSizeBytes)
            {
                Console.WriteLine("Warning: JPEG file size exceeds the acceptable maximum.");
            }
            else
            {
                Console.WriteLine("JPEG file size is within the acceptable range.");
            }
        }
    }
}