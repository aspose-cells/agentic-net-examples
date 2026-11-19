// Title: How to enable text wrapping for a TextBox shape in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set textbox IsTextWrapped property to true | wrap long text inside Excel textbox shape with Aspose.Cells .NET | enable multiline text in Aspose.Cells TextBox shape | C# Aspose.Cells example for text wrapping in worksheet textbox | prevent textbox overflow by enabling text wrap in Aspose.Cells workbook
// Tags: Aspose.Cells TextBox text wrapping | C# Excel shape IsTextWrapped property | Aspose.Cells set textbox multiline | Excel worksheet textbox formatting Aspose.Cells | Aspose.Cells workbook save with wrapped textbox

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a TextBox shape to the first worksheet, assigns a long sentence, enables automatic line breaking by setting IsTextWrapped to true, and saves the file as an .xlsx workbook.
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

            // Add a textbox shape to the worksheet.
            // Parameters: upper left row, upper left column, lower right row, lower right column, width, height
            // Width and height are in pixels.
            TextBox textbox = sheet.Shapes.AddTextBox(1, 1, 4, 4, 200, 80);

            // Set the text content of the textbox
            textbox.Text = "This is a very long sentence that should automatically wrap onto multiple lines within the textbox when text wrapping is enabled.";

            // Enable text wrapping so long sentences break onto multiple lines
            textbox.IsTextWrapped = true;

            // Optionally keep the textbox size fixed (AutoSize is not available in this version)
            // No action needed; the size remains as defined by width and height.

            // Define output file path
            string outputPath = "TextBoxWrap.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
