// Title: Apply an Outer Shadow Preset to Shape Text with Aspose.Cells (C#)
// Description: Creates a new workbook, adds a rectangle auto‑shape, sets its text and font, then uses TextOptions.Shadow.PresetType to apply an outer shadow (OffsetBottom) to the shape's characters before saving the file as XLSX.
// Keywords: Aspose.Cells outer shadow | C# shape text shadow | TextOptions.Shadow | PresetShadowType | Aspose.Cells shape formatting | Excel shape text shadow C# | Aspose.Cells TextOptions example
// Common Searches: Aspose.Cells set outer shadow on shape text | C# TextOptions.Shadow.PresetType sample | How to add shadow to shape text in Excel using Aspose.Cells | PresetShadowType OffsetBottom C# Aspose.Cells | Apply shadow to auto shape text Aspose.Cells
// Developer Intent: Add a predefined outer shadow effect to the text inside an Excel shape using Aspose.Cells for .NET.
// Use Cases: Design report headings with bold, colored text and a subtle shadow for visual emphasis. | Generate Excel templates where labels inside diagram shapes have a consistent shadow style. | Automate styling of workflow diagrams by applying a preset shadow to shape captions.
// AI Prompts: Show how to change the shadow preset to OffsetTopRight for shape text with Aspose.Cells. | Provide C# code to set a custom shadow color and blur radius on shape text using TextOptions.Shadow. | Explain how to list all PresetShadowType values and let a user choose one in an Aspose.Cells project.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;
using System.Drawing;

namespace AsposeCellsShadowExample
{
    // Creates a new workbook, adds a rectangle auto‑shape, sets its text and font, then uses TextOptions.Shadow.PresetType to apply an outer shadow (OffsetBottom) to the shape's characters before saving the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle auto shape to the worksheet
            // Parameters: shape type, upper left row, top, upper left column, left, height, width
            Shape shape = sheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 4, 4, 4, 4, 100, 700);
            shape.Fill.FillType = FillType.None; // No fill for clarity
            shape.Text = "Hello World !!!";

            // Get the FontSetting for the entire text of the shape
            FontSetting fontSetting = shape.Characters(0, "Hello World !!!".Length);
            // Access TextOptions to modify text formatting
            TextOptions textOptions = fontSetting.TextOptions;

            // Set basic text formatting (optional)
            textOptions.Name = "Calibri";
            textOptions.Size = 54;
            textOptions.IsBold = true;
            textOptions.Color = Color.Green;

            // Apply an outer shadow preset to the text
            // Using OffsetBottom as an example of an outer shadow type
            textOptions.Shadow.PresetType = PresetShadowType.OffsetBottom;

            // Save the workbook
            workbook.Save("ShadowOnTextShape.xlsx", SaveFormat.Xlsx);
        }
    }
}
