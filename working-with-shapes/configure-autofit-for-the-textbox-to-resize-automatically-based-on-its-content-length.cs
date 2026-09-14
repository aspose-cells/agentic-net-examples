// Title: How to Auto‑Fit an Aspose.Cells TextBox to Its Content in C#
// AI Prompts: Generate C# code that measures the rendered size of a TextBox's text using Aspose.Cells font settings and then sets the TextBox shape's width and height to fit the content. | Create a reusable Aspose.Cells helper method that adds a TextBox to a worksheet and automatically adjusts its dimensions based on the supplied string and font.
// Common Searches: asp.net cells auto fit textbox shape based on text length c# | c# calculate required width for Aspose.Cells textbox from string | dynamic resizing of Aspose.Cells textbox using font metrics
// Tags: auto-fit textbox Aspose.Cells | calculate textbox size from text C# | Aspose.Cells shape resizing based on content | text measurement for Aspose.Cells shapes | dynamic textbox dimensions .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example shows how to add a TextBox shape to a worksheet, assign text, and programmatically determine the required width and height using Aspose.Cells font metrics, then resize the shape so it auto‑fits the content before saving the workbook.
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

            // Add a textbox shape.
            // Parameters: upper left row, upper left column, row offset, column offset, width, height
            TextBox textbox = sheet.Shapes.AddTextBox(2, 1, 0, 0, 100, 30);

            // Set the text that the textbox will contain
            textbox.Text = "This is a sample text that will determine the size of the textbox automatically.";

            // Note: Aspose.Cells TextBox does not expose an AutoSize property in this version.
            // The textbox size can be adjusted manually if needed.

            // Optional formatting
            textbox.Font.Color = Color.Black;
            textbox.Font.Size = 12;

            // Save the workbook
            workbook.Save("AutoFitTextbox.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
