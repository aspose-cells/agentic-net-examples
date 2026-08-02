// Title: Extract a Worksheet Background Image with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, accesses a worksheet, reads its BackgroundImage byte array, verifies its presence, and writes the bytes to an image file (e.g., PNG). Includes a fallback message when no background is set.
// Keywords: Aspose.Cells | Worksheet.BackgroundImage | C# extract Excel background | save worksheet background image | export Excel sheet picture | .NET image bytes | Excel background extraction
// Common Searches: extract background picture from Excel worksheet using Aspose.Cells C# | save worksheet background as image file .NET | retrieve background image bytes from Aspose.Cells worksheet | export Excel sheet background PNG
// Developer Intent: Obtain the background picture of a worksheet and write it to an external image file.
// Use Cases: Create thumbnails of worksheet backgrounds for documentation or reporting. | Batch‑extract backgrounds from multiple sheets to build a web gallery. | Compare a worksheet’s background image against a reference image for validation.
// AI Prompts: Generate C# code with Aspose.Cells that extracts a worksheet’s background image and saves it as a PNG. | Add robust error handling for cases where the worksheet has no background image. | Detect the image format of the extracted bytes before choosing the file extension.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, accesses a worksheet, reads its BackgroundImage byte array, verifies its presence, and writes the bytes to an image file (e.g., PNG). Includes a fallback message when no background is set.
class ExtractWorksheetBackgroundImage
{
    static void Main()
    {
        // Path to the source Excel file
        string workbookPath = "input.xlsx";

        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook(workbookPath);

        // Access the first worksheet (or any specific worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the background image as a byte array (Worksheet.BackgroundImage property)
        byte[] backgroundBytes = worksheet.BackgroundImage;

        // Check if a background image exists
        if (backgroundBytes != null && backgroundBytes.Length > 0)
        {
            // Define the output image file path (you may change the extension based on actual image format)
            string outputImagePath = "worksheet_background.png";

            // Save the image bytes to a file (save rule)
            File.WriteAllBytes(outputImagePath, backgroundBytes);

            Console.WriteLine($"Background image extracted and saved to: {outputImagePath}");
        }
        else
        {
            Console.WriteLine("The worksheet does not contain a background image.");
        }
    }
}
