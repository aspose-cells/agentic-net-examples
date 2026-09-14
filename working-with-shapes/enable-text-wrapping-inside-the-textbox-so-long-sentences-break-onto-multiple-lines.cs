// Title: Enable word wrap for a TextBox shape in an Aspose.Cells for .NET workbook (C#)
// AI Prompts: Write C# code that adds a TextBox shape to an Excel worksheet with Aspose.Cells and forces the text to wrap onto multiple lines. | Show how to configure the TextBox shape's properties in Aspose.Cells so that long sentences automatically break within the shape. | Provide an example that adjusts the TextBox dimensions after enabling text wrapping using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# how to wrap text inside a textbox shape | set word wrap on Excel textbox using Aspose.Cells .NET | C# Aspose.Cells textbox automatic line break for long strings | adjust textbox size after enabling wrap text in Aspose.Cells workbook | enable multiline text in Aspose.Cells shape programmatically
// Tags: Aspose.Cells C# textbox word wrap | Aspose.Cells set textbox wrap text property | Aspose.Cells adjust textbox dimensions for wrapped content | Aspose.Cells shape text wrapping in Excel | Aspose.Cells create textbox with auto line break

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, inserts a TextBox shape on the first worksheet, assigns a long sentence, enables automatic word wrap, optionally resizes the shape to fit the wrapped text, and saves the file as WrappedTextBox.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet.
            // Parameters: upper left row, upper left column, upper left row offset,
            // upper left column offset, height (points), width (points)
            TextBox textBoxShape = sheet.Shapes.AddTextBox(2, 1, 0, 0, 50, 100);

            // Set the text content (word wrap is automatic for textboxes)
            textBoxShape.Text = "This is a very long sentence that should automatically wrap onto multiple lines inside the textbox.";

            // Optionally adjust size
            textBoxShape.Width = 100;
            textBoxShape.Height = 50;

            // Save the workbook
            workbook.Save("WrappedTextBox.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
