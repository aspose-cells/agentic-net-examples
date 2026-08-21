// Title: C# – Set Shape Text Character Spacing to 5 Points with Aspose.Cells TextOptions
// Description: Creates a new workbook, adds a textbox shape, assigns "Sample Text", uses TextOptions.Spacing to set a 5‑point character gap, and saves the file as CharacterSpacingDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | TextOptions.Spacing | character spacing | shape text formatting | textbox Aspose.Cells | Excel shape text | adjust text spacing | Excel automation
// Common Searches: Aspose.Cells set character spacing | TextOptions spacing property C# example | how to change shape text spacing in Excel with Aspose | C# Aspose.Cells shape text formatting | increase textbox letter spacing Aspose.Cells
// Developer Intent: Apply a 5‑point character spacing to the text of a shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Design report headings where wider letter spacing improves visual hierarchy. | Create diagram labels that require consistent spacing for readability. | Generate Excel templates that enforce a 5‑point spacing rule on all shape texts before export. | Produce PDF versions of spreadsheets with precise text spacing in shapes.
// AI Prompts: Write C# code that adds a textbox to an Aspose.Cells worksheet and sets its TextOptions.Spacing to 5 points. | Explain how TextOptions.Spacing influences shape text and how to apply it to different shape types in Aspose.Cells. | Provide a C# loop that iterates over every shape in a worksheet and sets each shape's character spacing to 5 points.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsCharacterSpacingDemo
{
    // Creates a new workbook, adds a textbox shape, assigns "Sample Text", uses TextOptions.Spacing to set a 5‑point character gap, and saves the file as CharacterSpacingDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a text box shape to the worksheet
            Shape textBox = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 200, 100);
            textBox.Text = "Sample Text";

            // Access the TextOptions of the shape and set character spacing to 5 points
            TextOptions textOptions = textBox.TextOptions;
            textOptions.Spacing = 5.0; // 5 points spacing between characters

            // Save the workbook
            workbook.Save("CharacterSpacingDemo.xlsx");
        }
    }
}
