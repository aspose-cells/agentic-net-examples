// Title: Extract ODS worksheet page background image and save as JPEG using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# extract page background picture from ODS workbook | Save ODS worksheet background image as JPEG with Aspose.Cells | How to get ImageBytes of ODS page background using reflection in .NET | Convert ODS page background graphic to JPG programmatically | Read ODS file page setup background image using Aspose.Cells API
// Tags: Aspose.Cells ODS background image extraction | C# export ODS background as JPEG | reflection access BackgroundPicture property Aspose.Cells | retrieve ImageBytes from worksheet background | ODS to JPEG conversion using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace OdsPageBackgroundExport
{
    // The example loads an ODS workbook with Aspose.Cells, uses reflection to access the worksheet's PageSetup BackgroundPicture object, extracts its ImageBytes, and writes the bytes to a JPEG file, handling missing files and absent background graphics gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input ODS and output JPEG files
            string odsPath = "input.ods";
            string jpegPath = "page_background.jpg";

            try
            {
                // Verify that the source ODS file exists
                if (!File.Exists(odsPath))
                {
                    Console.WriteLine($"Input file not found: '{odsPath}'.");
                    return;
                }

                // Load the ODS workbook
                using (Workbook workbook = new Workbook(odsPath))
                {
                    // Get the first worksheet (adjust index if needed)
                    Worksheet sheet = workbook.Worksheets[0];

                    // Attempt to retrieve the page background picture via reflection
                    object backgroundPicture = sheet.PageSetup.GetType()
                        .GetProperty("BackgroundPicture")?
                        .GetValue(sheet.PageSetup, null);

                    if (backgroundPicture != null)
                    {
                        // Try to get the image bytes from the picture object
                        var imageBytesProp = backgroundPicture.GetType()
                            .GetProperty("ImageBytes");

                        if (imageBytesProp != null)
                        {
                            byte[] imageBytes = imageBytesProp.GetValue(backgroundPicture) as byte[];
                            if (imageBytes != null && imageBytes.Length > 0)
                            {
                                // Save the image bytes as a JPEG file
                                File.WriteAllBytes(jpegPath, imageBytes);
                                Console.WriteLine($"Page background saved to '{jpegPath}'.");
                            }
                            else
                            {
                                Console.WriteLine("Background picture found but contains no image data.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Background picture does not expose ImageBytes property.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("No page background graphic found in the ODS file.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
