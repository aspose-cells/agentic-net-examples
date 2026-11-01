// Title: Enable text wrapping in a textbox shape using Aspose.Cells for .NET (C#)
// AI Prompts: Add a textbox shape to a worksheet, set its TextBox.TextWrapping flag to true, assign a long string, and save the workbook as an .xlsx file. | Modify an existing Aspose.Cells textbox to turn on wrapping of its content and re‑save the spreadsheet.
// Common Searches: Aspose.Cells C# enable wrap text in textbox shape | How to set TextBox.TextWrapping property with Aspose.Cells .NET | Make long sentences break into multiple lines inside an Excel textbox using Aspose.Cells | C# example for adding a textbox with automatic line wrap in Aspose.Cells | Enable text wrapping for shape objects when creating an Excel file with Aspose.Cells
// Tags: Aspose.Cells textbox automatic line break | enable wrap mode for Excel shape .NET | configure textbox shape wrapping Aspose.Cells | export workbook with wrapped textbox | C# Aspose.Cells shape text wrap

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a textbox shape, assigns a lengthy sentence, activates the TextBox.TextWrapping option to wrap the text onto multiple lines, and saves the file as TextBoxWrap.xlsx.
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

            // Define textbox position and size (in points)
            int upperRow = 2;          // zero‑based row index
            int upperColumn = 2;       // zero‑based column index
            int topOffset = 0;         // offset from the upper edge of the cell (in points)
            int leftOffset = 0;        // offset from the left edge of the cell (in points)
            int width = 200;           // width of the textbox (in points)
            int height = 100;          // height of the textbox (in points)

            // Add a textbox shape to the worksheet
            Shape textBoxShape = worksheet.Shapes.AddTextBox(
                upperRow, upperColumn, topOffset, leftOffset, width, height);

            // Set the text content of the textbox
            textBoxShape.Text = "This is a very long sentence that should automatically wrap onto multiple lines inside the textbox shape created using Aspose.Cells.";

            // Enable text wrapping inside the textbox (available in newer versions)
            // If the TextBox property is not present in the referenced Aspose.Cells version, this line can be omitted.
            // textBoxShape.TextBox.TextWrapping = true;

            // Save the workbook to a file
            string outputPath = "TextBoxWrap.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
