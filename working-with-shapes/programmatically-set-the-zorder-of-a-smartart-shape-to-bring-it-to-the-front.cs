// Title: How to bring a SmartArt shape to the front (Z‑order) with Aspose.Cells for .NET
// Description: Loads or creates an Excel workbook, finds the first SmartArtShape, calls SmartArtShape.ToFrontOrBack(1) to move it forward in the Z‑order, and saves the file.
// Keywords: Aspose.Cells | .NET | C# | SmartArt | Z-order | bring shape to front | ToFrontOrBack | Excel shape layering | shape ordering | Aspose.Cells SmartArt
// Common Searches: Aspose.Cells move SmartArt forward | C# set Z order of Excel shape | bring SmartArt to front using Aspose.Cells | SmartArt ToFrontOrBack example | change shape layering in Excel with Aspose.Cells
// Developer Intent: Adjust the Z‑order of a SmartArt shape so it appears above other worksheet objects.
// Use Cases: Update an existing workbook so a specific SmartArt diagram overlays charts or images. | Add new SmartArt to a generated report and ensure it renders on top of all other shapes. | Process multiple worksheets in a workbook, bringing every SmartArt shape forward for consistent visual hierarchy. | Combine Z‑order changes with further shape formatting (color, size) after reordering.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells and moves a selected SmartArt shape to the front. | Write a method that accepts a Worksheet object and brings all SmartArtShape objects to the front using ToFrontOrBack. | Explain the purpose of the integer parameter in SmartArtShape.ToFrontOrBack and how different values affect Z‑order.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtZOrder
{
    // Loads or creates an Excel workbook, finds the first SmartArtShape, calls SmartArtShape.ToFrontOrBack(1) to move it forward in the Z‑order, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to a template workbook that already contains a SmartArt shape.
                const string templatePath = "SmartArtTemplate.xlsx";

                Workbook workbook;

                if (File.Exists(templatePath))
                {
                    // Load the template workbook.
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    // If the template is missing, create a new workbook and add a placeholder shape.
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    // Add a simple rectangle as a fallback (cannot add SmartArt directly in older versions).
                    ws.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 300, 200);
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Locate the first SmartArt shape in the worksheet.
                SmartArtShape smartArt = null;
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape is SmartArtShape sa)
                    {
                        smartArt = sa;
                        break;
                    }
                }

                if (smartArt != null)
                {
                    // Bring the SmartArt shape one position forward in the Z‑order.
                    smartArt.ToFrontOrBack(1);
                }
                else
                {
                    Console.WriteLine("No SmartArt shape found in the worksheet.");
                }

                // Save the workbook.
                const string outputPath = "SmartArtZOrderDemo.xlsx";
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
