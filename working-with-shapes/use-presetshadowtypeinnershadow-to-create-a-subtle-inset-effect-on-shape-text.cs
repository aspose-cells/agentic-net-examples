// Title: Apply InsideCenter PresetShadow to Shape Text with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a rectangle auto‑shape, set its text, configure font properties via TextOptions, and apply a subtle inset shadow using PresetShadowType.InsideCenter before saving as an XLSX file.
// Keywords: Aspose.Cells C# inner shadow | PresetShadowType InsideCenter example | shape text shadow Aspose.Cells | auto shape text formatting | .NET Excel shape effects | text shadow preset Aspose
// Common Searches: Aspose.Cells add inner shadow to shape text C# | PresetShadowType examples for shape text | how to apply inset shadow to Excel shape text | C# Aspose.Cells text shadow on auto shape | set text shadow effect in Aspose.Cells workbook
// Developer Intent: Add a subtle inset shadow to the text inside an auto‑shape using Aspose.Cells for .NET.
// Use Cases: Enhance report visuals by giving shape text a recessed appearance. | Combine custom font styling with an InsideCenter shadow for consistent branding. | Generate Excel dashboards where highlighted shape text stands out without using external images.
// AI Prompts: Show how to replace InsideCenter with PresetShadowType.OuterShadow for shape text in Aspose.Cells C#. | Provide C# code that adds multiple shadow layers with different offsets and colors to a single shape. | Explain how to combine a text outline with an inner shadow on shape text using Aspose.Cells.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to create a workbook, add a rectangle auto‑shape, set its text, configure font properties via TextOptions, and apply a subtle inset shadow using PresetShadowType.InsideCenter before saving as an XLSX file.
class InnerShadowExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape that will contain the text
        // Parameters: upper left row, top offset, upper left column, left offset, height, width
        Shape shape = sheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 2, 0, 2, 0, 150, 400);
        shape.Text = "Inset Shadow";

        // Obtain the TextOptions for the shape's text
        // Characters(start, length) returns a FontSetting; its TextOptions let us modify text formatting
        TextOptions textOpts = shape.Characters(0, shape.Text.Length).TextOptions;

        // Set some basic font properties (optional)
        textOpts.Name = "Calibri";
        textOpts.Size = 36;
        textOpts.IsBold = true;
        textOpts.Color = Color.DarkBlue;

        // Apply an inner shadow effect to the text
        // InsideCenter creates a subtle inset shadow around the text
        textOpts.Shadow.PresetType = PresetShadowType.InsideCenter;

        // Save the workbook
        workbook.Save("InnerShadowShapeText.xlsx", SaveFormat.Xlsx);
    }
}
