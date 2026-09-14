// Title: Move the 'Watermark' shape to the back layer in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel file with Aspose.Cells, locate the shape named 'Watermark' on a worksheet, and set its ZOrderPosition to 0 so it appears behind the cells. | Write C# code that changes the Z-order of any named shape in a worksheet to the back layer using the Shape.ZOrderPosition property of Aspose.Cells. | Update an existing Aspose.Cells workbook to reposition a watermark shape behind all other objects by adjusting its ZOrderPosition.
// Common Searches: Aspose.Cells C# how to send a shape to the back of the worksheet | set ZOrderPosition of a watermark shape in Excel using Aspose.Cells .NET | move named shape behind cells in an Excel file with Aspose.Cells API | C# Aspose.Cells change shape layering order | place watermark shape behind data in Excel workbook programmatically
// Tags: shape ZOrderPosition Aspose.Cells | move watermark to back layer C# | adjust Excel shape layering .NET | retrieve named shape worksheet Aspose.Cells | set shape order behind cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, finds the shape named "Watermark" on the first worksheet, sets its ZOrderPosition to 0 to place it behind all cells, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the shape named "Watermark"
                Shape watermarkShape = worksheet.Shapes["Watermark"];

                // If the shape exists, move it to the back layer by setting its Z-order position
                if (watermarkShape != null)
                {
                    // Lower ZOrderPosition values are rendered behind higher values
                    watermarkShape.ZOrderPosition = 0;
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
