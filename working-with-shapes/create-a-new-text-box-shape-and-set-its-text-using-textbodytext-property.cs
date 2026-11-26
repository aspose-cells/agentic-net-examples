// Title: Add a rectangle shape as a text box and set its text using TextBody.Text in Aspose.Cells for .NET (C#)
// Common Searches: c# how to add a textbox shape to an Excel worksheet with Aspose.Cells | using TextBody.Text to assign text to a rectangle shape in Aspose.Cells | apply bold blue Arial font to a shape textbox in Aspose.Cells .NET | configure shape placement to move and size with cells in Aspose.Cells | export workbook that contains a textbox shape using Aspose.Cells
// Tags: add rectangle shape textbox Aspose.Cells | set shape TextBody.Text property C# | shape font formatting Aspose.Cells | placement type MoveAndSize Aspose.Cells | save workbook with textbox shape Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextBoxDemo
{
    // Demonstrates creating a new Workbook, adding a rectangle shape that acts as a text box at cell B2, setting its placement to MoveAndSize, assigning the text "Hello, Aspose.Cells!" via the TextBody.Text property, applying Arial 14 bold blue font, and saving the file as TextBoxShapeDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape that will act as a text box.
                // Parameters: upper left row, upper left column, upper left offset (pixels),
                // lower right row, lower right column, lower right offset (pixels)
                // Here we place the shape starting at cell B2 (row 1, column 1) with size 200x100 pixels.
                Shape textBoxShape = worksheet.Shapes.AddRectangle(
                    1, // upperLeftRow
                    1, // upperLeftColumn
                    0, // upperLeftPixel
                    5, // lowerRightRow
                    3, // lowerRightColumn
                    0  // lowerRightPixel
                );

                // Set the shape's placement so it moves and resizes with cells (optional)
                textBoxShape.Placement = PlacementType.MoveAndSize;

                // Set the text of the shape
                textBoxShape.TextBody.Text = "Hello, Aspose.Cells!";

                // Adjust text formatting using the shape's Font property
                textBoxShape.Font.Name = "Arial";
                textBoxShape.Font.Size = 14;
                textBoxShape.Font.IsBold = true;
                textBoxShape.Font.Color = Color.Blue;

                // Save the workbook to a file
                string outputPath = "TextBoxShapeDemo.xlsx";
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
