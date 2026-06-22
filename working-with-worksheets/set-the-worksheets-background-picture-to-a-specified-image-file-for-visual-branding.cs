using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetBackgroundDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the background image file
                string imagePath = "branding_background.png";

                // Verify that the image file exists before loading
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {Path.GetFullPath(imagePath)}");
                    return;
                }

                // Load the image file into a byte array
                byte[] imageData = File.ReadAllBytes(imagePath);

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the worksheet's background image
                worksheet.BackgroundImage = imageData;

                // Save the workbook with the background applied
                string outputPath = "WorkbookWithBackground.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Worksheet background image applied successfully. Saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}