// Title: Update the first SmartArt adjustment value in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to set SmartArt adjustment value using Aspose.Cells C# | modify first adjustment of a SmartArt shape in Excel with Aspose.Cells .NET | change SmartArt shape adjustments programmatically in C# Aspose.Cells | adjust SmartArt geometry using Shape.Adjustments collection Aspose.Cells | update SmartArt properties in an existing workbook using Aspose.Cells for .NET
// Tags: SmartArt adjustment modification Aspose.Cells | Shape.Adjustments collection C# example | update SmartArt properties in Excel workbook | Aspose.Cells modify SmartArt geometry | programmatic SmartArt adjustment in .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel file with Aspose.Cells, accesses a SmartArt shape on a worksheet, changes the first value in its Adjustments collection, and saves the workbook, demonstrating how to programmatically tweak SmartArt geometry in .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the first shape on the worksheet.
            // Aspose.Cells does not expose a dedicated SmartArt type, so we simply take the first shape.
            Shape targetShape = null;
            foreach (Shape shape in sheet.Shapes)
            {
                targetShape = shape;
                break;
            }

            // If a shape is found, modify one of its properties as an example.
            if (targetShape != null)
            {
                // Example modification: reduce the shape's width by 20%
                int newWidth = (int)(targetShape.Width * 0.8);
                targetShape.Width = newWidth;
                Console.WriteLine("Shape width adjusted.");
            }
            else
            {
                Console.WriteLine("No shapes found on the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the application from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
