// Title: Copy a shape's absolute position to another shape in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get shape top and left coordinates from worksheet | How to align two shapes to the same location with Aspose.Cells .NET | Set target shape position equal to source shape in an .xlsx file using Aspose.Cells API | Retrieve absolute position of a shape and apply it to another shape in C# | Programmatically copy shape location in Excel with Aspose.Cells
// Tags: Aspose.Cells shape coordinate transfer | C# align worksheet shapes | Excel shape top left positioning Aspose | copy shape location .xlsx Aspose.Cells | shape alignment using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, obtains two shapes by name, reads the source shape's Top and Left values, assigns those coordinates to the target shape to align them, and saves the updated file.
class ShapeAlignment
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access the shapes collection on the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Retrieve the source shape (replace with actual shape name)
            Shape sourceShape = shapes["SourceShapeName"];
            if (sourceShape == null)
            {
                Console.WriteLine("Source shape not found.");
                return;
            }

            // Retrieve the target shape (replace with actual shape name)
            Shape targetShape = shapes["TargetShapeName"];
            if (targetShape == null)
            {
                Console.WriteLine("Target shape not found.");
                return;
            }

            // Align the target shape to the same coordinates as the source shape
            targetShape.Top = sourceShape.Top;
            targetShape.Left = sourceShape.Left;

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
