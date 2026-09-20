// Title: Send the "Background" shape to the back by setting its Z‑order to the lowest level in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells to locate a shape named "Background" and set its ZOrderPosition to 0. | Generate a .NET snippet that moves the "Background" shape behind all other worksheet objects by adjusting its Z‑order. | Create a C# program that loads an existing .xlsx file, finds the "Background" shape, sends it to the back, and saves the workbook.
// Common Searches: Aspose.Cells C# how to move a shape to the back of a worksheet | set ZOrderPosition to 0 for a specific shape in an Excel file using Aspose.Cells | C# code to send background image behind other objects in an .xlsx with Aspose.Cells | change layering order of shapes in Aspose.Cells workbook programmatically
// Tags: Aspose.Cells shape ZOrderPosition .NET | C# send Excel shape to back | background shape layering Aspose.Cells | adjust shape order in worksheet using Aspose.Cells | move shape behind others Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, finds the shape named "Background" on the first worksheet, sets its ZOrderPosition to 0 to place it behind all other objects, and saves the updated file.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the shape named "Background"
            Shape backgroundShape = null;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape.Name == "Background")
                {
                    backgroundShape = shape;
                    break;
                }
            }

            // If the shape is found, send it to the back by setting the lowest Z‑order position
            if (backgroundShape != null)
            {
                // Setting ZOrderPosition to 0 moves the shape behind all others
                backgroundShape.ZOrderPosition = 0;
            }
            else
            {
                Console.WriteLine("Shape named \"Background\" was not found.");
            }

            // Save the workbook with the updated Z‑order
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
