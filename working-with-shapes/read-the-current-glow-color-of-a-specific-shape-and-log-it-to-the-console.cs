// Title: Read and log the glow effect color of a specific shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Create C# code that opens an .xlsx workbook with Aspose.Cells, verifies a shape’s Glow property, and writes the resulting System.Drawing.Color to the console. | Provide a method in C# using Aspose.Cells that takes a worksheet and shape index, returns the glow property’s color if it exists, and handles cases where the glow is absent. | Show how to access the Glow.Color of a Shape object in Aspose.Cells and output it as a readable string.
// Common Searches: aspocells c# retrieve glow color from shape in excel file | how to get shape glow using aspocells .net | example code to read glow of an Excel shape with aspocells
// Tags: aspocells shape glow extraction | c# read excel shape glow attribute | aspocells retrieve shape glow property | excel worksheet shape glow retrieval

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an .xlsx workbook with Aspose.Cells, accesses the first worksheet, ensures at least one shape exists, checks that the shape has a Glow effect, extracts the Glow.Color as a System.Drawing.Color, and writes the color value to the console, with error handling for missing files, shapes, or glow.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Retrieve the first shape (or replace with a specific index/name as required)
            Shape shape = sheet.Shapes[0];

            // Ensure the shape has a glow effect
            if (shape.Glow == null)
            {
                Console.WriteLine("The shape does not have a glow effect.");
                return;
            }

            // Obtain the glow effect's color (convert CellsColor to System.Drawing.Color)
            Color glowColor = shape.Glow.Color.Color;

            // Output the glow color to the console
            Console.WriteLine($"Glow color of the shape: {glowColor}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
