// Title: How to center-align text inside a textbox shape using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates an Excel workbook, adds a TextBox shape, sets its TextHorizontalAlignment to Center, and saves the file with Aspose.Cells. | Provide a concise example demonstrating how to apply horizontal center alignment to a TextBox shape's content using Aspose.Cells in a .NET application. | Generate a snippet that shows creating a textbox, assigning text, centering it, and exporting the workbook with Aspose.Cells for C#.
// Common Searches: Aspose.Cells C# set TextBox shape horizontal alignment to center in Excel workbook | C# Aspose.Cells example for centering text inside a textbox shape | How to use TextHorizontalAlignment property with Aspose.Cells TextBox in .NET
// Tags: Aspose.Cells TextHorizontalAlignment usage | center alignment for Excel textbox C# | Aspose.Cells TextBox shape creation | C# workbook save with centered textbox | rich text formatting textbox Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, adds a textbox shape to the first worksheet, sets its text to "Hello World", applies the TextHorizontalAlignment property with a Center value to horizontally align the text, and saves the workbook as CenteredTextInTextbox.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox shape (row, column, upper‑left row offset, upper‑left column offset, width, height)
            Shape shape = worksheet.Shapes.AddTextBox(2, 1, 0, 0, 300, 100);

            // Cast the shape to a TextBox to access text‑related properties
            TextBox textBox = (TextBox)shape;

            // Set the textbox text
            textBox.Text = "Hello World";

            // Center‑align the text inside the textbox
            textBox.TextHorizontalAlignment = TextAlignmentType.Center;

            // Define output file path
            string outputPath = "CenteredTextInTextbox.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
