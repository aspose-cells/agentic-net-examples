// Title: Insert a Picture into Excel with Fixed Height (200 pt) and Auto‑scaled Width using Aspose.Cells for .NET
// Description: Demonstrates how to add an image to the first worksheet, set its height to 200 points, lock the aspect ratio so the width adjusts automatically, and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells picture insertion | C# set image height points | lock aspect ratio Excel | add image to worksheet .NET | auto‑scale picture width
// Common Searches: Aspose.Cells add image with fixed height | how to keep picture proportions in Excel using C# | set picture height 200 points Aspose.Cells example | auto‑adjust image width after setting height Aspose.Cells
// Developer Intent: Add an image to a worksheet, enforce a 200‑point height, and preserve its original proportions.
// Use Cases: Standardizing logo size in automated financial reports. | Displaying product photos at a uniform height in catalog generators. | Embedding screenshots in dashboards without distortion.
// AI Prompts: Generate C# code with Aspose.Cells that inserts a PNG at cell B2, sets height to 150 pt, and locks the aspect ratio. | Show how to insert multiple pictures, each with a specific height and locked proportions, then export the workbook. | Provide a snippet that checks for an image file before adding it as a picture with a fixed height in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add an image to the first worksheet, set its height to 200 points, lock the aspect ratio so the width adjusts automatically, and save the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            string imagePath = "image.jpg";

            // Verify that the image file exists before adding it
            if (File.Exists(imagePath))
            {
                // Add a picture to the worksheet (top-left corner at row 1, column 1)
                int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Set the picture height to 200 points
                picture.HeightPt = 200;

                // Lock the aspect ratio so the width adjusts automatically
                picture.IsAspectRatioLocked = true;
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Save the workbook
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
