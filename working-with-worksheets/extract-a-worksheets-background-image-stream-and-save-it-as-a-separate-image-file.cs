// Title: Extract Worksheet Background Image with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, reads the worksheet's BackgroundImage byte array, verifies its presence, and writes the data to an image file (e.g., PNG or JPEG). Includes a fallback message when no background is set.
// Keywords: Aspose.Cells | C# | extract worksheet background image | save Excel background as image | Worksheet.BackgroundImage | byte array to file | export Excel sheet picture | background image PNG | background image JPEG
// Common Searches: Aspose.Cells get worksheet background image C# | save Excel sheet background to file | extract background picture from .xlsx using .NET | Worksheet.BackgroundImage example | convert worksheet background to PNG
// Developer Intent: Retrieve a worksheet's background picture and store it as a standalone image file.
// Use Cases: Create a reusable image of a sheet's design for documentation. | Archive original background graphics before modifying the workbook. | Display the worksheet background on a web page alongside exported data.
// AI Prompts: Generate C# code that extracts the BackgroundImage from a specified worksheet and saves it as a PNG, handling the case where no image exists. | Write a method that loops through all worksheets in a workbook, extracts each background image, and saves the files using the worksheet names as filenames. | Explain how to detect the original image format (PNG, JPEG, etc.) from the byte array returned by Worksheet.BackgroundImage.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, reads the worksheet's BackgroundImage byte array, verifies its presence, and writes the data to an image file (e.g., PNG or JPEG). Includes a fallback message when no background is set.
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
            Console.WriteLine("The worksheet does not contain a background image.");
        }
    }
}
