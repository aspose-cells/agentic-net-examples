// Title: Send a worksheet shape to the back of the Z‑order hierarchy using Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# example that uses Aspose.Cells to move a specific worksheet shape to the back of the Z‑order so the cells underneath become visible. | Generate code that retrieves the first shape on a worksheet and sets its Z‑order index to zero with the Aspose.Cells .NET API. | Show how to reorder shapes in an Excel file by sending a rectangle shape behind all other objects using Aspose.Cells in C#.
// Common Searches: asp.net aspose.cells move shape to back of z-order | c# set excel shape behind cells using aspose.cells | how to change shape layering in an Excel worksheet with aspose.cells .net | aspose.cells example for sending a shape to the back of the stack
// Tags: Aspose.Cells shape ZOrderPosition | C# send Excel shape to back | Aspose.Cells worksheet shape layering | Excel shape behind cells Aspose.Cells | Aspose.Cells Z-order manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The code loads (or creates) a workbook, ensures a shape exists on the first worksheet, retrieves the first shape, sets its Z‑order index to zero to move it behind all other objects, and saves the workbook, making underlying cells visible.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Load existing workbook or create a new one if the input file is missing.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    // Add a sample shape so the example can demonstrate Z‑order manipulation.
                    Worksheet ws = workbook.Worksheets[0];
                    Shape sampleShape = ws.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);
                    sampleShape.Name = "SampleShape";
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one shape to manipulate.
                if (worksheet.Shapes.Count > 0)
                {
                    Shape shape = worksheet.Shapes[0];
                    // Move the shape to the back of the Z‑order hierarchy.
                    shape.ZOrderPosition = 0;
                }
                else
                {
                    Console.WriteLine("No shapes found in the worksheet.");
                }

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
