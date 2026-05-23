using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsWordIntegration
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new Excel workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.20);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.80);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(2.50);

                // ---------- Define a named range ----------
                AsposeRange namedRange = sheet.Cells.CreateRange("A1:B4");
                namedRange.Name = "ProductsTable";

                // ---------- Convert the named range to an image ----------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    OnlyArea = true // render only the used area
                };

                // Render the range to a PNG byte array
                byte[] rangeImageBytes = namedRange.ToImage(imgOptions);

                // ---------- Save the image ----------
                string outputImagePath = Path.Combine(Environment.CurrentDirectory, "ProductsTable.png");
                try
                {
                    File.WriteAllBytes(outputImagePath, rangeImageBytes);
                    Console.WriteLine($"Image saved to: {outputImagePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save image: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}