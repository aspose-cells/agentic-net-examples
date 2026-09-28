// Title: Insert a PNG logo at coordinates (100,200) in a PDF generated from an Aspose.Cells workbook using C#
// AI Prompts: Add a free‑floating picture from a PNG file to the first worksheet, set its Left property to 100 and Top property to 200 points, then save the workbook as a PDF. | Create a reusable method that receives an image path and X/Y point values, inserts the image as a floating picture, and returns the generated PDF bytes. | Adjust the example to calculate coordinates in millimeters, apply them to picture.Left and picture.Top, and verify the logo appears at the expected spot in the exported PDF.
// Common Searches: C# Aspose.Cells how to place an image at exact X Y coordinates before exporting to PDF | set picture.Left and picture.Top properties in points when generating PDF with Aspose.Cells | free floating picture insertion in Aspose.Cells workbook C# example | export worksheet to PDF with logo positioned at (100,200) using Aspose.Cells
// Tags: Aspose.Cells picture placement using points | C# insert PNG logo into worksheet | export workbook to PDF with positioned image | floating picture insertion Aspose.Cells | absolute coordinate image positioning PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // This program creates a new workbook, inserts a PNG logo as a free‑floating picture positioned at X=100, Y=200 points, and saves the worksheet as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the logo image file
            string logoPath = "logo.png";

            // Add the logo image to the worksheet if the file exists
            if (File.Exists(logoPath))
            {
                using (FileStream imgStream = new FileStream(logoPath, FileMode.Open, FileAccess.Read))
                {
                    // Insert the picture at cell A1 (row 0, column 0)
                    int pictureIndex = sheet.Pictures.Add(0, 0, imgStream);
                    Picture picture = sheet.Pictures[pictureIndex];

                    // Set the picture to free‑floating so it can be positioned with absolute coordinates
                    picture.Placement = PlacementType.FreeFloating;

                    // Position the picture at the required coordinates (100, 200)
                    // Left and Top are measured in points (1 point = 1/72 inch)
                    picture.Left = 100; // X‑coordinate
                    picture.Top = 200;  // Y‑coordinate
                }
            }
            else
            {
                Console.WriteLine($"Warning: Logo file '{logoPath}' not found. Skipping image insertion.");
            }

            // Save the workbook as a PDF file
            string outputPath = "output.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
