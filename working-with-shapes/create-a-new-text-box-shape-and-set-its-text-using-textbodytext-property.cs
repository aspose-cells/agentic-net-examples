// Title: Create a TextBox shape in Aspose.Cells (.NET) and set its text with TextBody.Text
// Description: Shows how to add a TextBox shape to a worksheet using Aspose.Cells for .NET, assign the displayed string via the TextBody.Text property, and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells | .NET | C# | AddTextBox | TextBox shape | TextBody.Text | set textbox text | worksheet shape | Excel automation | Aspose.Cells example
// Common Searches: Aspose.Cells add textbox shape | Set textbox text using TextBody.Text in Aspose.Cells | TextBody.Text property C# example | How to create a shape with text in Aspose.Cells | C# Aspose.Cells TextBox TextBody usage
// Developer Intent: Add a TextBox shape to a worksheet and assign its displayed text via the TextBody.Text property.
// Use Cases: Insert a labeled comment box in an automatically generated financial report. | Add instructional annotations to a data‑dashboard workbook template. | Create a reusable Excel template that includes pre‑filled guidance inside a TextBox.
// AI Prompts: Generate C# code that adds multiple TextBox shapes with distinct TextBody.Text values in a loop using Aspose.Cells. | Provide an example that changes the font family, size, and color of a TextBox's TextBody after setting its text. | Explain how to position a TextBox relative to specific cell coordinates and then set its text with TextBody.Text.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Shows how to add a TextBox shape to a worksheet using Aspose.Cells for .NET, assign the displayed string via the TextBody.Text property, and save the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: topRow, top (pixel), leftColumn, left (pixel), height (pixel), width (pixel)
        TextBox textBox = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 100, 200);

        // Set the text of the textbox using the TextBody.Text property
        // TextBody returns a FontSettingCollection; its Text property holds the shape's text
        textBox.TextBody.Text = "This text is set via TextBody.Text";

        // Save the workbook
        workbook.Save("TextBoxTextBodyDemo.xlsx");
    }
}
