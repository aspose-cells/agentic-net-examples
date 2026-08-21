// Title: C# – Set Alternative Text for a TextBox Shape in Aspose.Cells (Accessibility)
// Description: Creates a workbook, adds a TextBox shape to the first worksheet, assigns descriptive AlternativeText for screen‑reader support, and saves the file as TextboxWithAltText.xlsx.
// Keywords: Aspose.Cells C# set textbox alt text | AlternativeText property Aspose.Cells | accessibility shapes Aspose.Cells | screen reader textbox Aspose.Cells | add alt description to Excel shape
// Common Searches: how to add alt text to a textbox in Aspose.Cells | Aspose.Cells alternative text for shapes | C# set accessibility text for Excel textbox | Aspose.Cells screen reader support
// Developer Intent: Apply the TextBox.AlternativeText property to provide a readable description for screen‑reader users.
// Use Cases: Provide a summary of chart data inside a textbox for visually impaired readers. | Label a form field in an auto‑generated report with accessible text. | Add instructional guidance to a textbox so screen readers can convey usage directions.
// AI Prompts: Write C# code using Aspose.Cells that inserts a TextBox and sets its AlternativeText to a custom string. | Explain how Aspose.Cells reads the AlternativeText property for screen readers and any known limitations. | Show a loop that iterates through all TextBox shapes in a workbook and assigns unique alternative text to each.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a TextBox shape to the first worksheet, assigns descriptive AlternativeText for screen‑reader support, and saves the file as TextboxWithAltText.xlsx.
class SetTextboxAlternativeText
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: upper left row, upper left column, top offset (pixels), left offset (pixels), height (pixels), width (pixels)
        TextBox textBox = (TextBox)worksheet.Shapes.AddTextBox(2, 2, 0, 0, 100, 200);

        // Set the alternative (alt) text for accessibility (screen readers)
        textBox.AlternativeText = "Summary of the data displayed in this textbox";

        // Save the workbook to a file
        workbook.Save("TextboxWithAltText.xlsx");
    }
}
