// Title: How to access a shape's TextBody and retrieve the ShapeTextAlignment object for text formatting with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# get ShapeTextAlignment from a textbox shape | retrieve and set text alignment of a shape's TextBody using Aspose.Cells .NET | C# code to read shape TextBody and change horizontal alignment in an Excel workbook | how to obtain the ShapeTextAlignment object for a shape in Aspose.Cells | set textbox shape text alignment via TextBody with Aspose.Cells for .NET
// Tags: Aspose.Cells shape TextBody alignment | retrieve ShapeTextAlignment C# | modify textbox text alignment Aspose.Cells | Excel shape text formatting .NET | Aspose.Cells get text alignment object

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads or creates an Excel workbook, ensures a textbox shape exists, accesses the shape's TextBody to obtain its ShapeTextAlignment object, sets the horizontal text alignment to center, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a workbook with one default worksheet
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape; add a textbox if none exist
            Shape shape;
            if (sheet.Shapes.Count > 0)
            {
                shape = sheet.Shapes[0];
            }
            else
            {
                // Add a textbox shape at row 0, column 0 with no offset, width 200, height 100
                shape = sheet.Shapes.AddTextBox(0, 0, 0, 0, 200, 100);
                shape.Text = "Sample Text";
            }

            // Set horizontal alignment of the shape's text to center
            shape.TextHorizontalAlignment = TextAlignmentType.Center;

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
