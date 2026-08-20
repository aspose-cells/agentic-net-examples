// Title: Clear Shape Text Formatting to Default with Aspise.Cells for .NET
// Description: Shows how to add a rectangle shape, apply custom text styling (red color, 14‑pt size, bold) and then revert all text formatting to the default appearance using the ClearFormatting (or TextBody.Clear) method in Aspose.Cells for .NET.
// Keywords: Aspose.Cells clear shape text formatting | Aspose.Cells reset shape font | ClearFormatting shape Aspose.Cells | Shape TextOptions reset .NET | Aspose.Cells TextBody.Clear | C# remove shape text style | default shape text Aspose.Cells | Aspose.Cells shape formatting API
// Common Searches: how to clear shape text formatting Aspose.Cells C# | reset shape font to default Aspose.Cells | ClearFormatting method for shape text Aspose.Cells | remove custom text style from worksheet shape | Aspose.Cells shape TextBody.Clear example
// Developer Intent: Reset a worksheet shape's text formatting to the default style.
// Use Cases: Reusing a shape after custom styling by clearing previous font settings. | Ensuring consistent text appearance across multiple shapes in a generated workbook. | Applying a theme‑based default style to shape text before adding dynamic content.
// AI Prompts: Generate C# code that adds a shape, sets custom TextOptions, then clears all formatting with ClearFormatting in Aspose.Cells. | Explain the difference between Shape.TextBody.Clear and ClearFormatting for resetting shape text styles. | Provide a step‑by‑step tutorial for reverting shape text to default formatting in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Shows how to add a rectangle shape, apply custom text styling (red color, 14‑pt size, bold) and then revert all text formatting to the default appearance using the ClearFormatting (or TextBody.Clear) method in Aspose.Cells for .NET.
class ResetShapeTextFormatting
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 0);
            shape.Text = "Sample Text";

            // Apply custom text formatting
            TextOptions textOpts = shape.TextOptions;
            textOpts.Color = Color.Red;
            textOpts.Size = 14;
            textOpts.IsBold = true;

            // Reset all text formatting to defaults by clearing font settings
            // In newer Aspose.Cells versions, TextBody itself is a FontSettingCollection
            shape.TextBody.Clear();

            // Save the workbook
            workbook.Save("ResetShapeTextFormatting.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
