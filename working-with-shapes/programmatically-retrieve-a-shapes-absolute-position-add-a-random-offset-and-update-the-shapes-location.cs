// Title: Read a shape's absolute position, add a random offset, and reposition it in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx workbook with Aspose.Cells, reads the Top and Left properties of the first shape on the first worksheet, adds a random offset between -20 and +20 points to each coordinate, updates the shape's position, and saves the workbook. | Generate a C# example that checks for shapes on a worksheet, retrieves their absolute coordinates, applies a random displacement, and persists the modified workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# move shape by random offset | Get shape top left coordinates in Excel using Aspose.Cells .NET | Programmatically change shape location in an existing .xlsx file with Aspose.Cells | How to apply random displacement to Excel drawing objects in C# | Update shape position based on current coordinates Aspose.Cells example
// Tags: Aspose.Cells shape reposition C# | retrieve shape absolute coordinates Aspose.Cells | random displacement for shape coordinates | modify shape Top Left properties .NET | Excel workbook shape movement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing .xlsx file, accesses the first worksheet, reads the Top and Left values of the first shape, adds a random offset between -20 and +20 points to each coordinate, updates the shape's position, and saves the workbook as a new file.
class ShapePositionUpdater
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Get the first shape (change index for a specific shape)
            Shape shape = sheet.Shapes[0];

            // Retrieve the shape's current absolute position (in points)
            // Top and Left are integer values
            int currentTop = shape.Top;
            int currentLeft = shape.Left;

            // Generate random integer offsets between -20 and +20 points
            Random rnd = new Random();
            int offsetTop = rnd.Next(-20, 21);   // inclusive lower bound, exclusive upper bound
            int offsetLeft = rnd.Next(-20, 21);

            // Apply the offsets to obtain the new position
            int newTop = currentTop + offsetTop;
            int newLeft = currentLeft + offsetLeft;

            // Update the shape's location
            shape.Top = newTop;
            shape.Left = newLeft;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
