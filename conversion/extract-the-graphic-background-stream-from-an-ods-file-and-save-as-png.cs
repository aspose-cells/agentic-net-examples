// Title: Extract a worksheet background image from an ODS file and save it as PNG using Aspose.Cells for .NET
// Common Searches: C# extract background image from OpenDocument Spreadsheet with Aspose.Cells | How to save ODS worksheet background as a PNG file in .NET | Retrieve worksheet background stream from an .ods file using Aspose.Cells API | Aspose.Cells get background image bytes from ODS worksheet | Convert ODS sheet background to PNG programmatically in C#
// Tags: Aspose.Cells ODS worksheet background extraction | save worksheet background as PNG .NET | background image bytes from OpenDocument spreadsheet | C# extract ODS sheet background image | export ODS background to PNG using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an ODS workbook with Aspose.Cells, accesses a worksheet, reads its BackgroundImage byte array, and writes the bytes directly to a PNG file, handling missing files and absent background images.
class Program
{
    static void Main()
    {
        const string inputPath = "input.ods";
        const string outputPath = "background.png";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the ODS workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or any specific worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the background image bytes of the worksheet
            byte[] backgroundBytes = worksheet.BackgroundImage;

            // Ensure there is a background image present
            if (backgroundBytes != null && backgroundBytes.Length > 0)
            {
                // Save the raw image bytes directly to a PNG file
                File.WriteAllBytes(outputPath, backgroundBytes);
                Console.WriteLine($"Background image saved to \"{outputPath}\".");
            }
            else
            {
                Console.WriteLine("No background image found in the worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
