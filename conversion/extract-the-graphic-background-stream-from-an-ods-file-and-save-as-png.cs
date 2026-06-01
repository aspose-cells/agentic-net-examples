using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsExamples
{
    public class ExtractOdsPageBackground
    {
        public static void Run()
        {
            try
            {
                // Path to the source ODS file
                string odsPath = "input.ods";

                // Verify that the source file exists
                if (!File.Exists(odsPath))
                {
                    Console.WriteLine($"Error: The file \"{odsPath}\" was not found.");
                    return;
                }

                // Load the ODS workbook
                Workbook workbook = new Workbook(odsPath);

                // Access the first worksheet (you can change the index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Get the ODS page background settings
                OdsPageBackground background = sheet.PageSetup.ODSPageBackground;

                // Verify that the background is of type Graphic and contains data
                if (background.Type == OdsPageBackgroundType.Graphic && background.GraphicData != null)
                {
                    // Retrieve the raw graphic data (byte array)
                    byte[] imageData = background.GraphicData;

                    // Define the output PNG file path
                    string pngPath = "background.png";

                    // Write the image data to a PNG file
                    File.WriteAllBytes(pngPath, imageData);

                    Console.WriteLine($"Graphic background extracted and saved to: {pngPath}");
                }
                else
                {
                    Console.WriteLine("The ODS file does not contain a graphic page background.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExtractOdsPageBackground.Run();
        }
    }
}