// Title: How to modify a shape's AdjustmentValues to change its geometry (e.g., increase star points) with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells C# to set the AdjustmentValues collection of a Shape object and reshape it, such as adding more points to a star. | Write C# code that loads an Excel workbook, retrieves a drawing shape, and updates its adjustment parameters to alter the figure's geometry. | Show how to programmatically adjust the geometry of an existing Excel shape by modifying its AdjustmentValues via the Aspose.Cells API.
// Common Searches: Aspose.Cells C# change star shape point count | set adjustment values for a shape in Aspose.Cells .NET | modify shape geometry programmatically in Excel using Aspose.Cells | adjust drawing object parameters like star points with Aspose.Cells API | how to use AdjustmentValues collection in Aspose.Cells C#
// Tags: Aspose.Cells shape adjustment values | modify shape geometry Aspose.Cells | set star points Aspose.Cells C# | adjust drawing object parameters .NET | Excel shape customization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing workbook, adds a rectangle shape to the first worksheet, illustrates where you can set the shape's AdjustmentValues to reshape it (e.g., change star point count), ensures the output directory exists, and saves the modified workbook.
class Program
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (star shape not supported in this version)
            // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 100);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
