// Title: Apply an Outer Preset Shadow to Shape Text with Aspose.Cells (C#)
// Description: Creates a workbook, inserts a rectangle auto‑shape, configures its text (Calibri, 54 pt, bold, green), and applies the OffsetBottom outer preset shadow via TextOptions.Shadow.PresetType before saving as XLSX.
// Keywords: Aspose.Cells | C# shape text shadow | TextOptions.Shadow | PresetShadowType | OffsetBottom | auto shape rectangle | Excel export .NET | no fill shape | green bold text
// Common Searches: Aspose.Cells set outer shadow on shape text C# | TextOptions.Shadow.PresetType example Aspose.Cells | OffsetBottom preset shadow for shape text .NET | Add shadow to rectangle shape text using Aspose.Cells | Change shape text shadow color Aspose.Cells C#
// Developer Intent: Add an outer preset shadow to the text of an auto‑shape in an Excel file using Aspose.Cells for .NET.
// Use Cases: Highlight a title inside a shape with green bold font and a subtle shadow for report headings. | Create callout shapes on dashboards where the shadow improves text legibility against busy backgrounds. | Generate automated worksheets that need visually distinct shape labels without using fill colors.
// AI Prompts: Show how to switch the shadow preset to OffsetTopRight for a shape's text using Aspose.Cells TextOptions.Shadow.PresetType in C#. | Provide C# code that sets a custom shadow color and blur radius for shape text with Aspose.Cells. | Explain how to read the current TextOptions.Shadow settings of an existing shape and modify them programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;
using System.Drawing;

// Creates a workbook, inserts a rectangle auto‑shape, configures its text (Calibri, 54 pt, bold, green), and applies the OffsetBottom outer preset shadow via TextOptions.Shadow.PresetType before saving as XLSX.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = sheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 4, 4, 4, 4, 100, 700);
        shape.Fill.FillType = FillType.None;          // No fill for the shape
        shape.Text = "Hello World !!!";               // Set the shape's text

        // Access the TextOptions of the shape's text
        TextOptions textOptions = shape.TextOptions;

        // Optional: configure basic font properties
        textOptions.Name = "Calibri";
        textOptions.Size = 54;
        textOptions.IsBold = true;
        textOptions.Color = Color.Green;

        // Apply an outer preset shadow to the text
        // OffsetBottom is one of the outer shadow types
        textOptions.Shadow.PresetType = PresetShadowType.OffsetBottom;

        // Save the workbook to a file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
