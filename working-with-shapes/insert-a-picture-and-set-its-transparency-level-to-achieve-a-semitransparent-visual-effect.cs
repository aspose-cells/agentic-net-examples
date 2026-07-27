// Title: Insert a picture with semi‑transparent effect using Aspose.Cells for .NET (C#)
// Description: This example creates a new Workbook, adds a PNG file to cell B3, accesses the Picture object, and sets its FillFormat.Transparency to 0.5, producing a semi‑transparent visual before saving the workbook as an .xlsx file.
// Keywords: Aspose.Cells picture transparency C# | insert image into Excel worksheet .NET | FillFormat Transparency property | C# set picture opacity Aspose | Excel watermark Aspose.Cells | overlay logo with opacity in Excel
// Common Searches: Aspose.Cells set picture opacity C# | how to make an inserted image transparent in Excel using Aspose | add semi transparent picture to worksheet Aspose.Cells | C# example for picture transparency in Excel | Aspose.Cells watermark image transparency
// Developer Intent: Add an image to a worksheet and apply a 50 % transparency level.
// Use Cases: Create a light‑weight watermark that does not hide cell data. | Overlay a corporate logo on a report while keeping the content readable. | Design a background image for a chart area with reduced opacity to improve visual hierarchy.
// AI Prompts: Generate C# code that inserts a JPEG into cell C5 and sets its transparency to 30 % with Aspose.Cells. | Explain how the FillFormat.Transparency property works for pictures and what values are accepted in Aspose.Cells for .NET. | Show how to modify the opacity of an existing picture in an Excel file using Aspose.Cells without re‑adding the image.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureTransparencyDemo
{
    // This example creates a new Workbook, adds a PNG file to cell B3, accesses the Picture object, and sets its FillFormat.Transparency to 0.5, producing a semi‑transparent visual before saving the workbook as an .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file to be inserted
                string imagePath = "sample.png";

                // Verify that the image file exists before adding it
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Please ensure the file exists.");
                    return;
                }

                // Add the picture to the worksheet at row 2, column 2 (zero‑based indices)
                int pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Set the picture's transparency to 50% (0.5)
                picture.FillFormat.Transparency = 0.5;

                // Save the workbook
                string outputPath = "PictureTransparencyDemo.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved to '{outputPath}' with picture transparency set to 0.5.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
