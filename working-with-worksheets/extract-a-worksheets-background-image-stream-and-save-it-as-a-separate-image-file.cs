using System;
using System.IO;
using Aspose.Cells;

class ExtractWorksheetBackgroundImage
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any other by index/name)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the background image as a byte array
        byte[] backgroundBytes = worksheet.BackgroundImage;

        // Check if a background image exists
        if (backgroundBytes != null && backgroundBytes.Length > 0)
        {
            // Save the image bytes to a file (choose appropriate extension, e.g., .png or .jpg)
            string outputImagePath = "worksheet_background.png";
            File.WriteAllBytes(outputImagePath, backgroundBytes);
            Console.WriteLine($"Background image extracted and saved to: {outputImagePath}");
        }
        else
        {
            Console.WriteLine("The worksheet does not have a background image.");
        }
    }
}