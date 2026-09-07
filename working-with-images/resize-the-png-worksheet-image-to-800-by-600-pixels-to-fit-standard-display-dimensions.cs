// Title: Resize the first PNG image on an Excel worksheet to 800 × 600 pixels with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that opens a workbook, finds the first PNG picture on the first worksheet, and sets its Width to 800 and Height to 600 pixels before saving. | Update the Aspose.Cells snippet to iterate over all pictures on a worksheet, detect PNG format, and resize only those images to 800 × 600 pixels. | Enhance the example with error handling that logs a warning when no PNG pictures are found and prevents the program from crashing.
// Common Searches: Aspose.Cells C# resize PNG picture in Excel worksheet to specific pixel dimensions | How to set picture width and height in an .xlsx file using Aspose.Cells .NET | Programmatically change size of embedded image in Excel with Aspose.Cells | Resize first image on Excel sheet to 800x600 using Aspose.Cells library | C# code to adjust dimensions of PNG objects in a workbook with Aspose.Cells
// Tags: Aspose.Cells resize PNG picture | worksheet picture width height Aspose.Cells | set embedded image dimensions Excel .NET | programmatic picture scaling in workbook | adjust PNG size on Excel worksheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The program opens input.xlsx, accesses the first worksheet, resizes the first picture on that sheet to 800 × 600 pixels, and saves the modified workbook as output.xlsx, including file existence verification and basic exception handling.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one picture on the worksheet
                if (worksheet.Pictures.Count > 0)
                {
                    // Get the first picture
                    Picture picture = worksheet.Pictures[0];

                    // Resize the picture to 800x600 pixels
                    picture.Width = 800;
                    picture.Height = 600;
                }
                else
                {
                    Console.WriteLine("No pictures found on the worksheet.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
