// Title: How to read a shape's line spacing with ShapeTextAlignment before updating its text in Aspose.Cells for .NET
// AI Prompts: Get the current LineSpacing value from a shape's ShapeTextAlignment, modify the shape's Text, and save the workbook using Aspose.Cells C#. | Write C# code that accesses Shape.TextAlignment.LineSpacing, stores the value, changes the shape's content, and preserves the original spacing. | Demonstrate retrieving and logging a shape's line spacing, then updating the shape text without altering formatting in an Excel file.
// Common Searches: Aspose.Cells C# read shape line spacing before text change | ShapeTextAlignment.LineSpacing property example .NET | preserve shape text formatting when updating shape in Excel using Aspose | how to get current line spacing of a textbox shape in Aspose.Cells | retrieve shape text alignment settings prior to modification Aspose.Cells
// Tags: extract shape line spacing Aspose.Cells | ShapeTextAlignment.LineSpacing C# | update shape text preserve formatting .NET | access shape text alignment properties Aspose.Cells | modify textbox shape line spacing Excel C# | Aspose.Cells shape formatting retrieval

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook, accesses the first shape on the first worksheet, reads its current line spacing via the ShapeTextAlignment object, updates the shape's text while keeping the original spacing, and saves the modified file.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the first worksheet.");
                return;
            }

            // Get the first shape
            Shape shape = sheet.Shapes[0];

            // Display the current text of the shape (if applicable)
            Console.WriteLine($"Current shape text: {shape.Text}");

            // Update the shape's text
            shape.Text = "Updated by Aspose.Cells";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
