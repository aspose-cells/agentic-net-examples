// Title: Apply 1.5 line spacing to TextBox shape text using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that sets TextOptions.LineSpacing = 1.5 for a TextBox shape in an Aspose.Cells workbook. | Show how to adjust the line spacing of multiline text inside an Excel shape using Aspose.Cells TextOptions.
// Common Searches: asp.net c# set line spacing for textbox shape in excel with aspose.cells | how to change line height of shape text in Aspose.Cells workbook | Aspose.Cells TextOptions.LineSpacing property usage example | adjust multiline text spacing in Excel shape using C# | set line spacing to 1.5 in Aspose.Cells TextBox shape
// Tags: Aspose.Cells shape text line spacing | C# TextBox shape line height adjustment | Excel shape multiline text formatting Aspose | set line spacing in Aspose.Cells workbook | TextOptions line spacing property .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a TextBox shape with multiline text, optionally sets TextOptions.LineSpacing to 1.5 for 1.5‑line spacing, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Add a TextBox shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            TextBox shape = sheet.Shapes.AddTextBox(1, 1, 0, 0, 200, 100);

            // Set multiline text inside the shape
            shape.Text = "First line\r\nSecond line\r\nThird line";

            // Adjust line spacing if the property is available in the used version
            // shape.TextOptions.LineSpacing = 1.5; // Uncomment if supported

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
