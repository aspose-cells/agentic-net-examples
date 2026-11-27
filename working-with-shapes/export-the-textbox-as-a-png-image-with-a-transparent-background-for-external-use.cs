// Title: Export an Excel textbox shape to a transparent PNG image using Aspose.Cells for .NET
// Common Searches: how to save an Excel textbox as a PNG with transparent background using Aspose.Cells C# | Aspose.Cells export textbox shape to PNG with transparency in .NET | C# code to convert Excel textbox to transparent PNG image | extract textbox from worksheet as PNG with alpha channel using Aspose.Cells | shape to image PNG export with transparent background Aspose.Cells example
// Tags: export textbox shape to transparent PNG Aspose.Cells | ImageOrPrintOptions Transparent property C# | AutoShapeType.TextSimple textbox detection Aspose.Cells | Shape.ToImage method PNG export | 300 DPI PNG export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace ExportTextboxTransparent
{
    // The sample loads an Excel workbook, finds the first textbox (AutoShapeType.TextSimple) on the first worksheet, configures ImageOrPrintOptions for PNG format with transparency and 300 DPI, and uses the Shape.ToImage method to export the textbox as a transparent PNG file.
    class Program
    {
        static void Main(string[] args)
        {
            const string sourceFile = "input.xlsx";
            const string outputImage = "textbox.png";

            try
            {
                // Verify that the source file exists to avoid FileNotFoundException.
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Source file '{sourceFile}' not found.");
                    return;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(sourceFile);

                // Get the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Find the first textbox shape.
                Shape textboxShape = null;
                foreach (Shape shape in sheet.Shapes)
                {
                    // In recent Aspose.Cells versions, textbox shapes are represented by AutoShapeType.TextSimple.
                    if (shape.AutoShapeType == AutoShapeType.TextSimple)
                    {
                        textboxShape = shape;
                        break;
                    }
                }

                if (textboxShape == null)
                {
                    Console.WriteLine("No textbox shape was found in the worksheet.");
                    return;
                }

                // Configure image export options.
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    Transparent = true,
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                // Export the textbox shape to a PNG file with a transparent background.
                textboxShape.ToImage(outputImage, imgOptions);

                Console.WriteLine($"Textbox exported successfully to '{outputImage}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during export:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
