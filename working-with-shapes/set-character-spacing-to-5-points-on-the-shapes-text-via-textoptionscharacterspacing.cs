// Title: How to set character spacing of shape text to 5 points using Aspose.Cells TextOptions in C#
// Common Searches: C# Aspose.Cells set shape text character spacing to 5 points | Aspose.Cells TextOptions spacing property example for rectangle shape | Adjust character spacing of Excel shape text with Aspose.Cells API | Set text options spacing on a shape in an Aspose.Cells workbook | Increase shape text character spacing using Aspose.Cells in C#
// Tags: Aspose.Cells TextOptions spacing API | adjust shape text spacing C# | rectangle shape character spacing Aspose.Cells | Excel shape text formatting Aspose.Cells | set character spacing points in workbook

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeCharacterSpacingDemo
{
    // Creates a workbook, adds a rectangle shape, assigns text, sets the shape's TextOptions.Spacing to 5 points, prints the value, and saves the file as an .xlsx workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape shape = worksheet.Shapes.AddRectangle(2, 1, 0, 0, 200, 100);

                // Set the text of the shape
                shape.Text = "Aspose.Cells Character Spacing Demo";

                // Set character spacing to 5 points using the shape's TextOptions
                shape.TextOptions.Spacing = 5.0;

                // Verify the setting by reading it back
                Console.WriteLine($"Character spacing set to: {shape.TextOptions.Spacing} points");

                // Save the workbook to a file
                string outputPath = "ShapeCharacterSpacingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
