// Title: Auto‑fit a TextBox shape to its content using Aspose.Cells for .NET (C#)
// Description: A C# example that creates an Excel workbook with Aspose.Cells, inserts a TextBox shape, assigns a long string, activates the TextBoxOptions.ResizeToFitText flag, calls FitToTextSize to recalculate dimensions, and saves the file. The shape expands automatically so the full text is visible.
// Keywords: Aspose.Cells | .NET | C# | TextBox auto size | ResizeToFitText | FitToTextSize | Excel textbox resize | shape sizing | Aspose.Cells TextBoxOptions | GitHub Aspose.Cells example
// Common Searches: Aspose.Cells auto size textbox | ResizeToFitText property C# example | FitToTextSize method usage | Make Excel textbox resize automatically with Aspose.Cells | Aspose.Cells shape size based on text
// Developer Intent: Adjust a TextBox shape so its dimensions automatically follow the length of its text.
// Use Cases: Generate comment boxes in reports that grow with variable‑length notes. | Create dynamic labels on dashboards that expand to fit data‑driven strings. | Design printable spreadsheets where annotations never truncate.
// AI Prompts: Show how to enable automatic resizing for a TextBox in Aspose.Cells and apply the change before saving the workbook. | Provide a C# loop that processes every TextBox on a worksheet, sets ResizeToFitText, and calls FitToTextSize. | Explain the role of TextBoxOptions.ResizeToFitText versus invoking FitToTextSize in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// A C# example that creates an Excel workbook with Aspose.Cells, inserts a TextBox shape, assigns a long string, activates the TextBoxOptions.ResizeToFitText flag, calls FitToTextSize to recalculate dimensions, and saves the file. The shape expands automatically so the full text is visible.
class TextBoxAutoFitDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        TextBox textBox = worksheet.Shapes.AddTextBox(1, 1, 100, 100, 200, 50);

        // Set the text that will determine the required size
        textBox.Text = "This is a sample text that will cause the textbox to resize automatically based on its content length.";

        // Enable auto‑fit so the shape resizes to fit the text
        textBox.TextBoxOptions.ResizeToFitText = true;

        // Apply the size adjustment immediately (optional, ensures dimensions are updated before saving)
        textBox.FitToTextSize();

        // Save the workbook
        workbook.Save("TextBoxAutoFitDemo.xlsx");
    }
}
