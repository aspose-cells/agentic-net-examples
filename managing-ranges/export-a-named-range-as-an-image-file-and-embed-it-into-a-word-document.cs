// Title: Export a Named Range to a High‑Resolution PNG and Embed It in Word with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, creates a named range (A1:B5), configures PNG conversion at 300 dpi using ImageOrPrintOptions, converts the range to a byte array, saves it as RangeImage.png, and shows how the resulting image can be inserted into a Word document with Aspose.Words.
// Keywords: Aspose.Cells | C# | export named range as image | Excel range to PNG | ImageOrPrintOptions 300 DPI | Aspose.Words embed image | Excel to Word snapshot | high‑resolution Excel picture
// Common Searches: Aspose.Cells export range to PNG C# | Convert Excel named range to image | Save Excel range as high‑resolution picture | Embed Excel range image into Word using Aspose | Set DPI for Excel range image Aspose.Cells
// Developer Intent: Create a PNG picture of a specific Excel range and make it ready for insertion into a Word document.
// Use Cases: Add a visual snapshot of a report section to a Word manual. | Email a high‑resolution image of a data table without sharing the whole workbook. | Generate graphics for PowerPoint slides directly from Excel data.
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, defines a named range, and saves it as a 300 dpi PNG. | Show how to load the PNG produced from a named range and insert it into a Word document using Aspose.Words. | Explain how to change ImageOrPrintOptions to output the range as JPEG, adjust dimensions, or apply a transparent background.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// Loads an Excel workbook, creates a named range (A1:B5), configures PNG conversion at 300 dpi using ImageOrPrintOptions, converts the range to a byte array, saves it as RangeImage.png, and shows how the resulting image can be inserted into a Word document with Aspose.Words.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string imagePath = "RangeImage.png";

            // Verify that the input Excel file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a range (A1:B5) and assign a name
            AsposeRange namedRange = worksheet.Cells.CreateRange("A1:B5");
            namedRange.Name = "MyRange";

            // Set image conversion options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Convert the named range to an image (byte array)
            byte[] imageData = namedRange.ToImage(imgOptions);

            // Ensure the directory for the image exists
            string imageDir = Path.GetDirectoryName(imagePath);
            if (!string.IsNullOrEmpty(imageDir) && !Directory.Exists(imageDir))
            {
                Directory.CreateDirectory(imageDir);
            }

            // Save the image to a file
            File.WriteAllBytes(imagePath, imageData);
            Console.WriteLine($"Range image saved to: {imagePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
