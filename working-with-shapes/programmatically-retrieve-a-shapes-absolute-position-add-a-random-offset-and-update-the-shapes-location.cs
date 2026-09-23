// Title: Read a shape’s absolute position, add a random offset, and update its location in an Excel file using Aspose.Cells for .NET
// AI Prompts: Get the Top and Left values of the first shape on a worksheet, calculate a random shift between -20 and 20 points, and assign the new coordinates to the shape. | Loop through all shapes on a worksheet, apply a random shift within a defined range to each shape’s Top and Left properties, and save the workbook. | Change the shift range to -50 to +50 points, update the shape positions accordingly, and write the result to a new Excel file.
// Common Searches: Aspose.Cells .NET get shape absolute position in points | C# move Excel shape by random offset using Aspose.Cells | How to change shape Top and Left properties programmatically with Aspose.Cells | Apply random displacement to multiple shapes in a worksheet Aspose.Cells | Save workbook after repositioning shapes Aspose.Cells C#
// Tags: Aspose.Cells retrieve shape coordinates | Aspose.Cells set shape top left | Aspose.Cells random shape displacement | Aspose.Cells move shapes worksheet | Aspose.Cells update shape position C# | Aspose.Cells save workbook after repositioning

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, reads the Top and Left values of the first shape on the first worksheet, adds a random offset between –20 and +20 points, updates the shape’s position, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Retrieve the first shape
            Shape shape = worksheet.Shapes[0];

            // Current absolute position (points)
            int currentTop = shape.Top;
            int currentLeft = shape.Left;

            // Random offset between -20 and +20 points
            Random rnd = new Random();
            int offsetTop = (int)(rnd.NextDouble() * 40 - 20);
            int offsetLeft = (int)(rnd.NextDouble() * 40 - 20);

            // Apply the random offset
            shape.Top = currentTop + offsetTop;
            shape.Left = currentLeft + offsetLeft;

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
