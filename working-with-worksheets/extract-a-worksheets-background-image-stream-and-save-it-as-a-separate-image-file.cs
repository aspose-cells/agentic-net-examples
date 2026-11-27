// Title: Extract and save a worksheet background image as PNG using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get worksheet background image bytes | How to save Excel worksheet background to PNG file with Aspose.Cells | Extract worksheet background image stream in .NET | Aspose.Cells export worksheet background picture to file | Read and write worksheet background image using MemoryStream in C#
// Tags: worksheet background image extraction Aspose.Cells | write background bytes to PNG in C# | memory stream copy of Excel background image | retrieve background picture bytes .NET | export worksheet background picture Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, accesses the first worksheet, extracts its BackgroundImage byte array, and saves the image directly as a PNG file. It also demonstrates writing the same bytes to a MemoryStream and then to a second PNG file, handling cases where the worksheet has no background image.
class ExtractWorksheetBackground
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputImagePath = "worksheet_background.png";
            const string outputStreamImagePath = "worksheet_background_stream.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the background image bytes (may be null)
            byte[] backgroundBytes = sheet.BackgroundImage;

            if (backgroundBytes != null && backgroundBytes.Length > 0)
            {
                // Save the background image directly to a file
                try
                {
                    File.WriteAllBytes(outputImagePath, backgroundBytes);
                    Console.WriteLine($"Background image saved to: {outputImagePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving image to file: {ex.Message}");
                }

                // Optionally, write the image bytes to a stream first and then to another file
                try
                {
                    using (MemoryStream ms = new MemoryStream(backgroundBytes))
                    using (FileStream file = new FileStream(outputStreamImagePath, FileMode.Create, FileAccess.Write))
                    {
                        ms.CopyTo(file);
                    }
                    Console.WriteLine($"Background image saved via stream to: {outputStreamImagePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving via stream: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("The selected worksheet does not have a background image.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
