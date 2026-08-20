// Title: C# – Send a Shape to the Back and Verify Z‑Order with Aspose.Cells for .NET
// Description: This example creates a workbook, adds two overlapping rectangle shapes, moves the second shape to the back using ToBack() or ToFrontOrBack(-1), prints each shape's ZOrderPosition to confirm the layering, ensures the output folder exists, and saves the file as an XLSX workbook.
// Keywords: Aspose.Cells shape Z-order | move shape to back Aspose.Cells | ToBack method C# | ToFrontOrBack negative Aspose.Cells | verify ZOrderPosition | Excel shape layering .NET | Aspose.Cells drawing order
// Common Searches: how to send a shape to the back in Aspose.Cells | Aspose.Cells check ZOrderPosition after reordering | C# Aspose.Cells ToFrontOrBack example | move Excel shape behind other objects using Aspose | set large negative Z-order Aspose.Cells
// Developer Intent: Move a specific shape behind all other drawing objects in a worksheet and confirm its Z‑order value.
// Use Cases: Place a background image behind charts and text in generated reports. | Keep a watermark shape hidden behind data cells. | Control the visual stacking of overlapping charts, pictures, or text boxes.
// AI Prompts: Generate C# code that uses Aspose.Cells to send a shape to the back by setting a large negative Z‑order and then output its ZOrderPosition. | Show an Aspose.Cells example that adds multiple shapes, moves one to the back with ToBack() or ToFrontOrBack(-1), and saves the workbook. | Explain how Aspose.Cells calculates ZOrderPosition values after calling ToFrontOrBack with a negative index.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // This example creates a workbook, adds two overlapping rectangle shapes, moves the second shape to the back using ToBack() or ToFrontOrBack(-1), prints each shape's ZOrderPosition to confirm the layering, ensures the output folder exists, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add two overlapping rectangles for visual verification
                Shape rect1 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
                Shape rect2 = worksheet.Shapes.AddRectangle(30, 30, 100, 100, 0, 0);

                // Send rect2 to the back using the proper API (move one position back)
                rect2.ToFrontOrBack(-1); // or rect2.ToBack();

                // Verify Z-order positions (lower value means farther back)
                Console.WriteLine("rect1 ZOrderPosition: " + rect1.ZOrderPosition);
                Console.WriteLine("rect2 ZOrderPosition (after ToFrontOrBack): " + rect2.ZOrderPosition);

                // Define output file path
                string outputPath = "ZOrderDemo.xlsx";

                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
