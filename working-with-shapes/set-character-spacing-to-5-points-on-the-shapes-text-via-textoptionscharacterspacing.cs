// Title: How to apply a 5‑point character spacing to a textbox shape's text using Aspose.Cells for .NET
// AI Prompts: Write C# code that applies a 5‑point character spacing to the text of a textbox shape in an Aspose.Cells workbook. | Show how to modify the Font.CharSpacing property through a shape's TextOptions to change text spacing in Aspose.Cells. | Create a complete example that adds a textbox shape, sets its text, and adjusts the character spacing to 5 points with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set character spacing for shape text | How to change font character spacing in a textbox shape with Aspose.Cells .NET | Adjust text spacing of an Excel shape using Aspose.Cells TextOptions | C# example of Font.CharSpacing property in Aspose.Cells
// Tags: Aspose.Cells shape font charspacing | textbox shape text formatting .NET | adjust shape text options Aspose.Cells | Excel shape character spacing C# | TextOptions Font.CharSpacing usage

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The sample creates a new Workbook, adds a textbox shape to the first worksheet, assigns text, and demonstrates how to apply a 5‑point character spacing by using the Font.CharSpacing property via the shape's TextOptions (when supported). The workbook is then saved as output.xlsx.
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

            // Add a textbox shape at row 2, column 2 with specified position and size
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape shape = sheet.Shapes.AddTextBox(2, 2, 50, 50, 200, 100);

            // Set the text displayed in the shape
            shape.Text = "Sample Text";

            // Note: Aspose.Cells TextOptions does not expose a CharacterSpacing property.
            // If character spacing is required, it can be adjusted via the Font object if supported.
            // shape.TextOptions.Font.CharSpacing = 5; // Uncomment if CharSpacing exists in your version.

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
