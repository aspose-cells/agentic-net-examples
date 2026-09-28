// Title: Bring a specific shape to the front of the Z‑order stack in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Retrieve a shape by its name and assign ZOrderPosition = 0 with Aspose.Cells in C#. | Programmatically move a rectangle shape to the top of the Z‑order in an Excel worksheet using the Aspose.Cells .NET API. | Change the layering of a targeted shape in a workbook by updating its ZOrderPosition property via C# code.
// Common Searches: Aspose.Cells C# bring shape to front of Z-order | set ZOrderPosition for a shape in an Excel file using .NET | how to change shape layering order in an Aspose.Cells workbook | move Excel shape to top layer programmatically C# | retrieve shape by name and adjust Z-order Aspose.Cells
// Tags: Aspose.Cells shape ZOrderPosition .NET | C# set Excel shape front layer | adjust shape layering Aspose.Cells | move rectangle shape to top in workbook | Excel shape Z-order manipulation with Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads or creates a workbook, obtains a shape named "MyShape", sets its ZOrderPosition to 0 to place it at the front of the Z‑order stack, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook or create a new one if the file is missing
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    // Add a sample shape named "MyShape"
                    Shape shape = ws.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);
                    shape.Name = "MyShape";
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the shape by its name
                Shape targetShape = worksheet.Shapes["MyShape"];
                if (targetShape != null)
                {
                    try
                    {
                        // Bring the shape to the front by setting its Z‑order position to the topmost (0)
                        targetShape.ZOrderPosition = 0;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to adjust Z‑order: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("Shape 'MyShape' not found.");
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
