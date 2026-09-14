// Title: Insert a PNG picture into an Excel worksheet at cell A1, set its height to 200 points, and keep the aspect ratio using Aspose.Cells for .NET
// AI Prompts: Create C# code that loads a PNG file, places it at cell A1 of a new workbook, locks its aspect ratio, and sets the picture height to 200 points using Aspose.Cells. | Provide a C# example that streams an image into an Excel sheet, enforces aspect‑ratio preservation, and adjusts the picture height to 200 points with the Aspose.Cells API.
// Common Searches: Aspose.Cells C# add image to specific cell and define height in points | How to preserve image proportions when inserting a picture into an Excel file with Aspose.Cells | Set picture height to 200 points while keeping aspect ratio in Aspose.Cells workbook | Programmatically insert a PNG into an Excel worksheet using Aspose.Cells .NET
// Tags: add picture to worksheet Aspose.Cells C# | set picture height points Aspose.Cells | lock aspect ratio image Aspose.Cells | insert PNG into Excel workbook Aspose.Cells | picture size adjustment Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// Creates a new workbook, inserts a PNG image at cell A1, locks its aspect ratio, sets the picture height to 200 points, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Path to the image you want to insert
            string imagePath = "image.png";

            // Verify that the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Insert the picture at cell A1 (row 0, column 0)
            using (FileStream imgStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                // Add returns the index of the newly added picture
                int pictureIndex = sheet.Pictures.Add(0, 0, imgStream);
                var picture = sheet.Pictures[pictureIndex];

                // Lock the aspect ratio so width adjusts automatically when height changes
                picture.IsLockAspectRatio = true;

                // Set the picture height to 200 points (1 point = 1/72 inch)
                picture.Height = 200;
            }

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
