using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper row, upper column, lower row, lower column, width, height
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1, 1,                     // upper row, upper column
                1, 1,                     // lower row, lower column
                150, 40);                 // width, height

            // Set the custom text that will appear inside the shape
            shape.Text = "Additional Resources";

            // Add a hyperlink that points to an external website
            shape.AddHyperlink("https://www.example.com/resources");

            // Optional: format the text to look like a hyperlink
            shape.Font.Color = Color.Blue;
            shape.Font.Underline = FontUnderlineType.Single;

            // Save the workbook to a file
            string outputPath = "ShapeWithHyperlink.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}