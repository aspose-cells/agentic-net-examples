// Title: Aspose.Cells .NET – Move a SmartArt Shape to the Front (Z‑Order)
// Description: This example shows how to load an existing or new workbook, get the first worksheet, locate the first SmartArt object, and call Shape.ToFrontOrBack(1) to raise its Z‑order so it appears above all other shapes before saving the file.
// Keywords: Aspose.Cells SmartArt Z-order | C# bring shape to front | Shape.ToFrontOrBack Aspose.Cells | Excel shape ordering .NET | move SmartArt forward
// Common Searches: Aspose.Cells move SmartArt to front | C# set Z‑order of Excel shape | How to bring a SmartArt shape forward with Aspose.Cells | Shape.ToFrontOrBack usage example
// Developer Intent: Adjust the Z‑order of a SmartArt shape so it renders on top of other worksheet objects.
// Use Cases: Ensure a SmartArt diagram overlays charts and images in automated report generation. | Correct the visual hierarchy of a template workbook before distribution. | Programmatically prioritize SmartArt when multiple shapes are added to a worksheet.
// AI Prompts: Generate C# code with Aspose.Cells that moves the first SmartArt shape to the front of a worksheet. | Create a snippet that iterates all shapes in a worksheet and sends every SmartArt object to the front using ToFrontOrBack. | Explain the ToFrontOrBack method parameters and how they affect shape Z‑order in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // This example shows how to load an existing or new workbook, get the first worksheet, locate the first SmartArt object, and call Shape.ToFrontOrBack(1) to raise its Z‑order so it appears above all other shapes before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Load existing workbook or create a new one if the file is missing
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    workbook.Save(inputPath);
                }

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Locate the first SmartArt shape and bring it to the front
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape.IsSmartArt)
                    {
                        shape.ToFrontOrBack(1); // move forward in Z‑order
                        break; // assume only one SmartArt needs adjustment
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
