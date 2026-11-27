// Title: Add a PNG background image to a worksheet from a byte stream and save the workbook as ODS using Aspose.Cells for .NET
// Common Searches: how to add a background image to an Excel worksheet from a byte array using Aspose.Cells C# | save a workbook as ODS with a custom worksheet background picture in Aspose.Cells | set worksheet background image from a file stream and export to ODS format in .NET | Aspose.Cells C# assign PNG background to ODS page | load background image bytes and apply to worksheet in Aspose.Cells
// Tags: worksheet background image byte array Aspose.Cells | save workbook as ODS with background picture | assign graphic stream to ODS page background | Aspose.Cells set worksheet background C# | export ODS with custom background Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example reads a PNG file into a byte array, assigns it to the worksheet's BackgroundImage property, and then saves the workbook as an ODS file with the custom background using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            string imagePath = "background.png";

            // Ensure the image file exists before attempting to read it
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Load the image bytes and set it as the worksheet background
            byte[] imageBytes = File.ReadAllBytes(imagePath);
            sheet.BackgroundImage = imageBytes;

            // Save the workbook as ODS
            string outputPath = "output.ods";
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
