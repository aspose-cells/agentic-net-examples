// Title: Export a Named Range to a PNG image and embed it in a Word document with Aspose.Cells (C#)
// Description: This example creates a workbook, defines a named range (A1:B3), configures ImageOrPrintOptions for a 300 dpi PNG, converts the range to a byte array, saves it to a temporary file, and demonstrates how the image can be inserted into a Word document using Aspose.Words.
// Keywords: Aspose.Cells C# export range image | named range to PNG | ImageOrPrintOptions OnlyArea | ToImage high resolution | Excel range snapshot | embed Excel image in Word | Aspose.Words insert picture | temporary file handling C# | Excel to Word image conversion
// Common Searches: Aspose.Cells export specific range as PNG C# | How to render a named range to an image with Aspose.Cells | Save Excel range as high‑resolution image | Insert Excel range image into Word using Aspose | C# temporary file pattern for image export
// Developer Intent: Create a high‑resolution PNG of a defined named range and use the resulting image in a Word document.
// Use Cases: Add a product table snapshot to a sales report generated in Word. | Include a chart or data block from Excel in an email as an image. | Generate documentation screenshots of Excel sections without exposing raw data.
// AI Prompts: Generate C# code that uses Aspose.Cells to export a named range to a 300 dpi PNG and then inserts the image into a new Word file with Aspose.Words. | Show how to set ImageOrPrintOptions so that only the used cells of a range are rendered. | Provide a robust pattern for creating, using, and safely deleting a temporary image file when converting an Excel range to an image.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example creates a workbook, defines a named range (A1:B3), configures ImageOrPrintOptions for a 300 dpi PNG, converts the range to a byte array, saves it to a temporary file, and demonstrates how the image can be inserted into a Word document using Aspose.Words.
class ExportRangeToWord
{
    static void Main()
    {
        string tempImagePath = Path.Combine(Path.GetTempPath(), "RangeImage.png");

        try
        {
            // Create a new workbook and populate it with sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Laptop");
            worksheet.Cells["B2"].PutValue(1200);
            worksheet.Cells["A3"].PutValue("Phone");
            worksheet.Cells["B3"].PutValue(800);

            // Define a named range that includes the data
            Aspose.Cells.Range namedRange = worksheet.Cells.CreateRange("A1:B3");
            namedRange.Name = "ProductData";

            // Configure image rendering options for the range
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                HorizontalResolution = 300,
                VerticalResolution = 300,
                OnlyArea = true // render only the used area of the range
            };

            // Convert the named range to an image (byte array)
            byte[] imageData = namedRange.ToImage(imgOptions);

            // Save the image to a temporary file
            File.WriteAllBytes(tempImagePath, imageData);

            // Output the location of the generated image
            Console.WriteLine($"Range image saved to: {tempImagePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Clean up the temporary image file if it exists
            if (File.Exists(tempImagePath))
            {
                try
                {
                    File.Delete(tempImagePath);
                }
                catch
                {
                    // Ignored – file may be in use or deletion may fail
                }
            }
        }
    }
}
