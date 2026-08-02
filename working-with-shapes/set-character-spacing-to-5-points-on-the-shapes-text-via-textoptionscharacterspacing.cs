// Title: C# – Set Shape Text Character Spacing to 5 Points with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a rectangle shape, assign text, and use TextOptions.Spacing to set the character spacing to 5 points before saving the Excel file.
// Keywords: Aspose.Cells C# shape text spacing | TextOptions.Spacing property | set character spacing Excel shape | adjust shape text spacing Aspose | Excel shape formatting C# | character spacing 5 points | Aspose.Cells shape text formatting
// Common Searches: Aspose.Cells set character spacing for shape text | C# TextOptions.Spacing example | how to change spacing of text inside Excel shape | set shape text spacing Aspose.Cells .NET | increase character spacing rectangle shape Excel
// Developer Intent: Apply a 5‑point character spacing to the text of a shape using Aspose.Cells for .NET.
// Use Cases: Design banner shapes with spaced lettering for clearer headings in automated reports. | Create Excel templates where shape labels need consistent readability through uniform spacing. | Batch‑style multiple worksheet shapes by applying the same character spacing value programmatically.
// AI Prompts: Generate C# code that sets a custom character spacing for shape text with Aspose.Cells. | Explain the units and effect of the TextOptions.Spacing property in Aspose.Cells. | Show how to loop through all shapes on a worksheet and apply identical character spacing.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to create a workbook, add a rectangle shape, assign text, and use TextOptions.Spacing to set the character spacing to 5 points before saving the Excel file.
class SetCharacterSpacing
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);
        shape.Text = "Sample Text";

        // Access the TextOptions of the shape
        TextOptions textOptions = shape.TextOptions;

        // Set character spacing to 5 points
        textOptions.Spacing = 5.0;

        // Save the workbook
        workbook.Save("CharacterSpacingDemo.xlsx");
    }
}
