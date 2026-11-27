// Title: Extract page background image from an ODS workbook and save it as a JPEG using Aspose.Cells for .NET
// Common Searches: how to get the page background picture from an ODS file with Aspose.Cells in C# | C# code to extract ODS worksheet background image and write to JPEG | using Aspose.Cells to read ODS page background graphic and save as jpg | retrieve ImageBytes of ODS background picture via reflection Aspose.Cells | save ODS page background as separate image file using .NET
// Tags: Aspose.Cells ODS background extraction | write background picture to JPEG C# | PageSetup.BackgroundPicture via reflection | ImageBytes property usage Aspose.Cells | ODS workbook page background to image

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an ODS workbook, uses reflection to access the worksheet's PageSetup.BackgroundPicture, reads the picture's ImageBytes, and writes the raw bytes to a JPEG file.
class OdsPageBackgroundExtractor
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.ods";
            const string outputFile = "output.jpg";

            // Verify that the input ODS file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the ODS workbook
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Use reflection to safely access PageSetup.BackgroundPicture (may not exist in older versions)
            PropertyInfo bgPicProp = sheet.PageSetup.GetType().GetProperty("BackgroundPicture", BindingFlags.Public | BindingFlags.Instance);
            if (bgPicProp == null)
            {
                Console.WriteLine("BackgroundPicture property is not available in this Aspose.Cells version.");
                return;
            }

            // Retrieve the Picture object
            Picture background = bgPicProp.GetValue(sheet.PageSetup) as Picture;
            if (background == null)
            {
                Console.WriteLine("No page background image found in the ODS file.");
                return;
            }

            // Use reflection to get ImageBytes from the Picture object
            PropertyInfo imgBytesProp = background.GetType().GetProperty("ImageBytes", BindingFlags.Public | BindingFlags.Instance);
            if (imgBytesProp == null)
            {
                Console.WriteLine("ImageBytes property is not available on the Picture object.");
                return;
            }

            byte[] imageBytes = imgBytesProp.GetValue(background) as byte[];
            if (imageBytes == null || imageBytes.Length == 0)
            {
                Console.WriteLine("Background picture does not contain image data.");
                return;
            }

            // Save the raw image bytes directly to a file
            File.WriteAllBytes(outputFile, imageBytes);
            Console.WriteLine($"Page background saved as {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
