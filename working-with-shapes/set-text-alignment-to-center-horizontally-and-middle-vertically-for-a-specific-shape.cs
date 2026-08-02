// Title: Center Shape Text Horizontally & Vertically with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a rectangle shape, assigns the text "Centered Text", and uses the TextHorizontalAlignment and TextVerticalAlignment properties to center the text both horizontally and vertically before saving the file.
// Keywords: Aspose.Cells | C# shape text alignment | TextHorizontalAlignment | TextVerticalAlignment | center text in Excel shape | rectangle shape Aspose.Cells | Excel shape formatting .NET
// Common Searches: Aspose.Cells center text in shape C# | Set vertical alignment for shape text Aspose.Cells | C# align shape caption in Excel | Aspose.Cells shape text alignment example
// Developer Intent: Apply centered horizontal and vertical alignment to the text of a specific shape using Aspose.Cells for .NET.
// Use Cases: Create a report header with a title centered inside a rectangle shape. | Design dashboard widgets where labels need to be perfectly centered within shapes. | Generate Excel templates with multiple shapes that share uniform centered captions.
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle shape and centers its text horizontally and vertically. | Explain the effect of TextHorizontalAlignment and TextVerticalAlignment on shape text in Aspose.Cells. | Show how to change shape text alignment to left/top or right/bottom using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a rectangle shape, assigns the text "Centered Text", and uses the TextHorizontalAlignment and TextVerticalAlignment properties to center the text both horizontally and vertically before saving the file.
    public class ShapeTextAlignmentDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape (parameters: upper left row, upper left column,
                // upper left row offset, upper left column offset, width, height)
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);

                // Set the shape's text
                shape.Text = "Centered Text";

                // Align text horizontally to center
                shape.TextHorizontalAlignment = TextAlignmentType.Center;

                // Align text vertically to middle (center)
                shape.TextVerticalAlignment = TextAlignmentType.Center;

                // Save the workbook
                string outputPath = "CenteredShape.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShapeTextAlignmentDemo.Run();
        }
    }
}
