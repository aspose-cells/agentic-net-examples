// Title: Auto‑fit TextBox shape in Aspose.Cells for .NET – resize automatically to content
// Description: Shows how to create a workbook, add a TextBox shape, set its text, enable auto‑fit with TextBoxOptions.ResizeToFitText, optionally invoke FitToTextSize for immediate sizing, and save the Excel file.
// Keywords: Aspose.Cells | .NET | C# | TextBox auto fit | ResizeToFitText | FitToTextSize | auto resize shape | Excel textbox size | Aspose.Cells example
// Common Searches: Aspose.Cells auto fit textbox | Resize textbox to fit text in .NET | TextBoxOptions ResizeToFitText usage | FitToTextSize method Aspose.Cells | How to make Excel textbox auto‑size with Aspose
// Developer Intent: Enable a TextBox shape to automatically adjust its width and height to the length of its text in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate comment boxes that expand to show full remarks without clipping. | Create dynamic labels that adapt their dimensions to varying data lengths. | Design templates where notes are inserted as auto‑sizing textboxes, preserving layout consistency.
// AI Prompts: Write C# code with Aspose.Cells to add a textbox that auto‑fits its content and apply a custom font and background color. | Show how to iterate over a list of strings and place an auto‑fit textbox for each item on successive rows. | Explain the difference between manually calculating textbox dimensions and using ResizeToFitText with FitToTextSize.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Shows how to create a workbook, add a TextBox shape, set its text, enable auto‑fit with TextBoxOptions.ResizeToFitText, optionally invoke FitToTextSize for immediate sizing, and save the Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        TextBox textBox = worksheet.Shapes.AddTextBox(1, 1, 0, 0, 200, 50);
        textBox.Text = "This is a sample text that will cause the textbox to automatically resize based on its content length.";

        // Enable auto‑fit so the shape resizes to fit the text
        textBox.TextBoxOptions.ResizeToFitText = true;

        // Recalculate the size immediately (optional, ensures the shape reflects the new setting)
        textBox.FitToTextSize();

        // Save the workbook
        workbook.Save("AutoFitTextBoxDemo.xlsx");
    }
}
