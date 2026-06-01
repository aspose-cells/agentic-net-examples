using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsExamples
{
    public class OdsPageBackgroundToJpeg
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Path to the source ODS file
                string odsPath = "input.ods";

                // Path for the output JPEG file
                string jpegPath = "background.jpg";

                // Verify that the input file exists
                if (!File.Exists(odsPath))
                {
                    Console.WriteLine($"Input file not found: {odsPath}");
                    return;
                }

                // Load the ODS workbook
                Workbook workbook = new Workbook(odsPath);

                // Access the ODS page background of the first worksheet
                OdsPageBackground background = workbook.Worksheets[0].PageSetup.ODSPageBackground;

                if (background == null)
                {
                    Console.WriteLine("No ODS page background available.");
                    return;
                }

                // Retrieve the graphic data (raw image bytes)
                byte[] graphicData = background.GraphicData;

                if (graphicData == null || graphicData.Length == 0)
                {
                    Console.WriteLine("No graphic background found in the ODS file.");
                    return;
                }

                // Save the raw image bytes directly as a JPEG file
                File.WriteAllBytes(jpegPath, graphicData);

                Console.WriteLine($"Graphic background extracted and saved to: {jpegPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}