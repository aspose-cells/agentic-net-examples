// Title: Get the absolute X and Y coordinates of a named shape (Logo) in an Excel file using Aspose.Cells for .NET
// Common Searches: how to read the X and Y position of a shape named Logo with Aspose.Cells in C# | Aspose.Cells retrieve absolute coordinates of a worksheet shape | C# code to log shape location in an Excel workbook using Aspose | find shape position in points using Aspose.Cells API | debugging shape placement coordinates in .xlsx with Aspose.Cells
// Tags: Aspose.Cells shape position extraction | C# read shape X Y values | worksheet shape named lookup Aspose | Excel shape coordinate logging | Aspose.Cells shape placement debugging

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook, searches the first worksheet for a shape named "Logo", reads its X and Y properties (absolute coordinates in points), and writes the coordinates to the console while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify the appropriate one)
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the shape named "Logo"
            Shape logoShape = null;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape.Name == "Logo")
                {
                    logoShape = shape;
                    break;
                }
            }

            if (logoShape != null)
            {
                // Retrieve absolute X and Y coordinates (in points)
                double x = logoShape.X;
                double y = logoShape.Y;

                // Log the coordinates for debugging
                Console.WriteLine($"Logo shape position - X: {x}, Y: {y}");
            }
            else
            {
                Console.WriteLine("Shape 'Logo' not found.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
