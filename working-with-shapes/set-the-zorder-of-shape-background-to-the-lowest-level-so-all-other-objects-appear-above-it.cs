// Title: Send a background shape to the back of an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set background rectangle behind other shapes | How to change ZOrderPosition of a shape in an Excel file with Aspose.Cells | Move a shape to the back layer in a .NET Excel workbook using Aspose.Cells | C# code to send a shape to the back in an Aspose.Cells worksheet | Set shape Z-order to lowest value in Aspose.Cells for .NET
// Tags: Aspose.Cells shape ZOrderPosition C# | background shape layering Excel .NET | shape backmost order Aspose.Cells | Excel worksheet shape ordering example | modify shape Z-order Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsZOrderDemo
{
    // The example loads or creates an Excel workbook, adds a rectangle named "Background" and two foreground shapes, finds the shape named "Background", sets its ZOrderPosition to 0 (the backmost layer), and saves the workbook, demonstrating how to send a background shape to the back using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output file paths.
            // If a path is supplied as the first argument, use it; otherwise create a demo workbook.
            string inputPath = args.Length > 0 ? args[0] : "DemoWorkbook.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "DemoWorkbook_ZOrdered.xlsx";

            try
            {
                Workbook workbook;
                Worksheet worksheet;

                if (File.Exists(inputPath))
                {
                    // Load existing workbook.
                    workbook = new Workbook(inputPath);
                    worksheet = workbook.Worksheets[0];
                    Console.WriteLine($"Loaded workbook \"{inputPath}\".");
                }
                else
                {
                    // Create a new workbook with sample shapes for demonstration.
                    workbook = new Workbook();
                    worksheet = workbook.Worksheets[0];

                    // Add a background shape (a rectangle filled with light gray).
                    Shape background = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 500, 500);
                    background.Name = "Background";
                    background.Fill.FillType = FillType.Solid;
                    background.Fill.SolidFill.Color = Color.LightGray;
                    background.HasLine = false; // Hide border line

                    // Add a few foreground shapes.
                    Shape rect1 = worksheet.Shapes.AddRectangle(50, 50, 0, 0, 150, 150);
                    rect1.Name = "Rect1";
                    rect1.Fill.FillType = FillType.Solid;
                    rect1.Fill.SolidFill.Color = Color.LightBlue;

                    Shape oval1 = worksheet.Shapes.AddOval(200, 200, 0, 0, 150, 100);
                    oval1.Name = "Oval1";
                    oval1.Fill.FillType = FillType.Solid;
                    oval1.Fill.SolidFill.Color = Color.LightCoral;

                    Console.WriteLine("Created demo workbook with a background shape and two foreground shapes.");
                }

                // Locate the shape named "Background".
                Shape backgroundShape = null;
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (string.Equals(shape.Name, "Background", StringComparison.OrdinalIgnoreCase))
                    {
                        backgroundShape = shape;
                        break;
                    }
                }

                if (backgroundShape == null)
                {
                    Console.WriteLine("No shape named \"Background\" was found. Exiting without changes.");
                    return;
                }

                // Set the Z-order to the lowest possible value (backmost).
                backgroundShape.ZOrderPosition = 0;

                Console.WriteLine($"Shape \"{backgroundShape.Name}\" Z-order set to {backgroundShape.ZOrderPosition} (backmost).");

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
