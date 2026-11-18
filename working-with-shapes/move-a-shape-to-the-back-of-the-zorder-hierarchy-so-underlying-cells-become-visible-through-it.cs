// Title: How to send an Excel shape to the back of the Z‑order using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to send shape behind cells | C# move Excel rectangle shape to back layer with Aspose.Cells | Make cells visible through a shape using Aspose.Cells .NET | Programmatically change Z‑order of a shape in an Aspose.Cells workbook | Aspose.Cells example: send worksheet shape to back of Z‑order
// Tags: Aspose.Cells shape Z-order API | C# Excel shape back positioning | Aspose.Cells worksheet shape layering | Programmatic shape Z-order adjustment .NET | Excel cell visibility through shape Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads or creates an Excel workbook, ensures a rectangle shape exists on the first worksheet, retrieves the shape, sets its ZOrderPosition to 0 to move it to the back of the Z‑order so underlying cells become visible, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one.
                if (File.Exists(inputFile))
                {
                    workbook = new Workbook(inputFile);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one worksheet

                    // Add a sample shape so the example works without an input file.
                    Worksheet wsInit = workbook.Worksheets[0];
                    Shape demoShape = wsInit.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);
                    demoShape.Name = "DemoShape";
                }

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one shape to manipulate.
                if (worksheet.Shapes.Count > 0)
                {
                    try
                    {
                        // Retrieve the first shape (or use a named shape).
                        Shape shape = worksheet.Shapes[0]; // or worksheet.Shapes["MyShape"]

                        // Move the shape to the back of the Z‑order by setting its ZOrderPosition to the minimum.
                        shape.ZOrderPosition = 0;
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Shape manipulation error: {exShape.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("No shapes found in the worksheet.");
                }

                // Save the modified workbook.
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
