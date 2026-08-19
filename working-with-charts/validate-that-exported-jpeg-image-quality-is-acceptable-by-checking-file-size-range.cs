// Title: Validate JPEG Export Quality in Aspose.Cells (C#) by Checking File Size Range
// Description: This C# sample builds a workbook, adds data, renders the first worksheet to a JPEG using ImageOrPrintOptions (Quality = 85), reads the resulting file size, and confirms it lies between 5 KB and 500 KB, indicating whether the image meets the required quality level.
// Keywords: Aspose.Cells JPEG export | C# image rendering options | check JPEG file size | Excel to JPEG quality validation | ImageOrPrintOptions Quality | Aspose.Cells chart export | file size range verification
// Common Searches: Aspose.Cells how to verify JPEG quality | C# check size of exported JPEG from Excel | Set JPEG quality in Aspose.Cells .NET | Validate exported image size Aspose.Cells | Determine acceptable JPEG file size after rendering worksheet
// Developer Intent: Ensure that a JPEG generated from an Excel worksheet with Aspose.Cells satisfies a predefined quality threshold by measuring its file size.
// Use Cases: Export a financial dashboard to JPEG with a specific quality setting and confirm the file stays within email attachment limits. | Integrate a size‑based quality gate into a CI pipeline that renders charts to JPEG and fails the build if the output exceeds defined bounds. | Create thumbnail previews of Excel reports and automatically reject images that are too large for storage quotas.
// AI Prompts: Write a C# method that accepts a worksheet, JPEG quality, and min/max size values, renders the sheet to JPEG with Aspose.Cells, and returns true if the file size is within the range. | Provide code to iterate over all worksheets in a workbook, export each as a JPEG with quality 90, and log any files whose size falls outside 10 KB‑200 KB. | Explain how to balance ImageOrPrintOptions Quality and DPI settings to achieve a target JPEG file size when using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsImageQualityCheck
{
    // This C# sample builds a workbook, adds data, renders the first worksheet to a JPEG using ImageOrPrintOptions (Quality = 85), reads the resulting file size, and confirms it lies between 5 KB and 500 KB, indicating whether the image meets the required quality level.
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
                ImageType = ImageType.Jpeg,
                Quality = 85 // Desired JPEG quality (0-100)
            };

            // Render the first worksheet page to a JPEG file
            string jpegPath = "ExportedSheet.jpg";
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, jpegPath);

            // Validate the exported JPEG file size is within an acceptable range
            const long minSizeBytes = 5_000;   // Example minimum size (5 KB)
            const long maxSizeBytes = 500_000; // Example maximum size (500 KB)

            FileInfo fileInfo = new FileInfo(jpegPath);
            long fileSize = fileInfo.Length;

            Console.WriteLine($"Exported JPEG size: {fileSize} bytes");

            if (fileSize >= minSizeBytes && fileSize <= maxSizeBytes)
            {
                Console.WriteLine("JPEG image quality is acceptable (file size within range).");
            }
            else
            {
                Console.WriteLine("JPEG image quality is NOT acceptable (file size out of range).");
            }
        }
    }
}
