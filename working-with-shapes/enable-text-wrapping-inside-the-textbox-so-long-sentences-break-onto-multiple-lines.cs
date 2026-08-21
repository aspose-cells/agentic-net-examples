// Title: Enable Automatic Line Wrapping in a TextBox Shape with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a TextBox shape to the first worksheet, inserts a long sentence, activates WrapTextInShape, disables overflow, and saves the file as TextboxWrapDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells TextBox wrap | WrapTextInShape property | TextBoxOptions AllowTextToOverflow false | C# Aspose.Cells textbox example | shape text wrapping Aspose
// Common Searches: how to wrap text inside a textbox shape Aspose.Cells | C# enable WrapTextInShape Aspose.Cells | prevent text overflow in Aspose.Cells textbox | multiline textbox Aspose.Cells .NET
// Developer Intent: Activate line wrapping for long strings inside a TextBox shape and stop the text from spilling outside the shape.
// Use Cases: Generating reports where comments need to fit inside fixed‑size label shapes. | Building dashboards with multiline annotations that stay within shape borders. | Creating invoices where address or note fields are placed in a textbox that must wrap automatically.
// AI Prompts: Provide C# code that enables text wrapping for a TextBox shape and sets a custom font using Aspose.Cells. | Show how to auto‑size a textbox after turning on WrapTextInShape in Aspose.Cells. | Explain the difference between TextBoxOptions.WrapTextInShape and TextBoxOptions.AllowTextToOverflow.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// This example creates a workbook, adds a TextBox shape to the first worksheet, inserts a long sentence, activates WrapTextInShape, disables overflow, and saves the file as TextboxWrapDemo.xlsx using Aspose.Cells for .NET.
class EnableTextboxWrapping
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a TextBox shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        TextBox textBox = sheet.Shapes.AddTextBox(1, 0, 1, 0, 200, 100);

        // Set a long text that needs wrapping
        textBox.Text = "This is a very long sentence that should automatically wrap inside the textbox shape when wrapping is enabled.";

        // Enable text wrapping within the shape
        textBox.TextBoxOptions.WrapTextInShape = true;

        // Ensure text does not overflow the shape boundaries
        textBox.TextBoxOptions.AllowTextToOverflow = false;

        // Save the workbook
        workbook.Save("TextboxWrapDemo.xlsx");
    }
}
