// Title: How to center text horizontally and vertically in a shape’s TextBody using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set horizontal alignment for textbox shape | center vertical text in Excel shape with Aspose.Cells API | retrieve and modify ShapeTextAlignment object in a .NET workbook | format shape TextBody alignment using Aspose.Cells for Excel | C# example for aligning text inside a shape in Aspose.Cells
// Tags: Aspose.Cells shape text alignment | C# Aspose.Cells TextHorizontalAlignment property | Aspose.Cells TextVerticalAlignment usage | Excel textbox alignment Aspose.Cells | Aspose.Cells ShapeTextAlignment object

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads or creates an Excel workbook, ensures a textbox shape exists, then sets the shape's TextHorizontalAlignment and TextVerticalAlignment properties to Center, effectively centering the text within the shape's TextBody, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet wsNew = workbook.Worksheets[0];

                // Add a textbox shape so we have something to format.
                // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
                Shape txtShape = wsNew.Shapes.AddTextBox(1, 0, 0, 0, 200, 100);
                txtShape.Text = "Sample Text";
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Ensure at least one shape is present.
            if (sheet.Shapes.Count == 0)
            {
                Shape txtShape = sheet.Shapes.AddTextBox(1, 0, 0, 0, 200, 100);
                txtShape.Text = "Sample Text";
            }

            Shape shape = sheet.Shapes[0];

            // Set horizontal and vertical alignment to Center.
            shape.TextHorizontalAlignment = TextAlignmentType.Center;
            shape.TextVerticalAlignment = TextAlignmentType.Center;

            // Save the modified workbook.
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
