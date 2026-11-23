// Title: Reposition an Excel worksheet shape with custom top/left offsets using Aspose.Cells for .NET
// Common Searches: c# aspose.cells shift shape position by specific points | how to get and set shape top left coordinates in Aspose.Cells workbook | add margin to a picture shape in Excel using Aspose.Cells .NET | adjust worksheet shape location programmatically with Aspose.Cells | retrieve shape absolute position and apply offset in Aspose.Cells C#
// Tags: shape top left offset Aspose.Cells | reposition excel shape programmatically C# | apply margin offset to worksheet shape | Aspose.Cells shape positioning points | modify shape coordinates in .NET workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel file, reads the first shape's Top and Left properties (in points), adds defined margin offsets, updates the shape's absolute position, and saves the workbook.
class ShapePositionAdjuster
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify the desired one)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains shapes before accessing
            if (sheet.Shapes.Count > 0)
            {
                // Get the first shape
                Shape shape = sheet.Shapes[0];

                // Current absolute position (in points)
                int currentTop = shape.Top;
                int currentLeft = shape.Left;

                // Margin offsets (in points)
                int offsetTop = 10;   // move down by 10 points
                int offsetLeft = 15;  // move right by 15 points

                // New position (cast to int if needed)
                shape.Top = currentTop + offsetTop;
                shape.Left = currentLeft + offsetLeft;
            }
            else
            {
                Console.WriteLine("No shapes found on the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
