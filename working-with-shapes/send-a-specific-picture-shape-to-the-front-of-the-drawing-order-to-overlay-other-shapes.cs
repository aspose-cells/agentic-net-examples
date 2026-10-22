// Title: How to bring a picture shape to the front of the drawing order in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a PNG image into cell B2 of a worksheet and set its ZOrderPosition to the highest index using Aspose.Cells in C#. | Reorder an existing picture shape to the top of the drawing stack in an Aspose.Cells workbook and save the file. | Load or create an Excel workbook, add a picture, adjust its Z-order so it overlays all other shapes, then export as XLSX with Aspose.Cells.
// Common Searches: Aspose.Cells C# bring picture shape to front of other shapes | set ZOrderPosition for picture in Aspose.Cells worksheet | how to change drawing order of images in Excel using Aspose.Cells .NET | move Excel picture to top layer programmatically with Aspose.Cells | C# Aspose.Cells picture overlay other worksheet shapes
// Tags: Aspose.Cells picture ZOrderPosition | C# adjust drawing order of worksheet shapes | Excel image layering with Aspose.Cells | Aspose.Cells shape front layering | C# add image to worksheet cell Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads (or creates) an Excel workbook, inserts a PNG picture into cell B2, sets its ZOrderPosition to the total picture count to move it to the front of all shapes, and saves the workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";
                string imagePath = "image.png";

                // Load existing workbook or create a new one if the file does not exist
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Ensure there is at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    workbook.Worksheets.Add();
                }

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add picture only if the image file exists
                if (File.Exists(imagePath))
                {
                    // Add picture at cell B2 (row index 1, column index 1) and retrieve its index
                    int pictureIndex = sheet.Pictures.Add(1, 1, imagePath);

                    // Get the picture object
                    Picture picture = sheet.Pictures[pictureIndex];

                    // Bring the picture to the front by setting a high Z-order position
                    picture.ZOrderPosition = sheet.Pictures.Count;
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the updated drawing order
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
