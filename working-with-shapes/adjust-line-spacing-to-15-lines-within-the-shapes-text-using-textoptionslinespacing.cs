// Title: Set 1.5 line spacing for text inside a rectangle shape using Aspose.Cells TextParagraph in C#
// Common Searches: C# Aspose.Cells how to change line spacing of shape text | set 1.5 line spacing for rectangle shape text with Aspose.Cells | adjust TextParagraph line spacing percentage in Aspose.Cells C# | modify line spacing of shape TextBody using Aspose.Cells API | increase line spacing inside shape text Aspose.Cells example
// Tags: Aspose.Cells shape line spacing API | TextParagraph line spacing percentage | rectangle shape text formatting Aspose.Cells | TextBody line spacing property | configure line spacing for shape text

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// The example creates a workbook, adds a rectangle shape with multiline text, iterates through each TextParagraph in the shape's TextBody, sets LineSpaceSizeType to Percentage and LineSpace to 150 (1.5 line spacing), and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape that will contain the text
            // Parameters: upper left row, upper left row offset, upper left column, upper left column offset, width, height
            Shape shape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 300, 150);
            shape.Text = "First line\nSecond line\nThird line";

            // Adjust line spacing for each paragraph (each line is a separate paragraph)
            TextParagraphCollection paragraphs = shape.TextBody.TextParagraphs;
            foreach (TextParagraph paragraph in paragraphs)
            {
                // Use percentage unit: 100 = single spacing, 150 = 1.5 line spacing
                paragraph.LineSpaceSizeType = LineSpaceSizeType.Percentage;
                paragraph.LineSpace = 150;
            }

            // Save the workbook to a file
            string outputPath = "ShapeLineSpacingDemo.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}
