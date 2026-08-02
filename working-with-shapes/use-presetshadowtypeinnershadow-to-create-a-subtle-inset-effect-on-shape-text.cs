// Title: Create an Inset Shadow for Shape Text with PresetShadowType.InsideCenter in Aspose.Cells for .NET
// Description: Shows how to add a rectangle shape to an Excel workbook, set its text font, and apply a subtle inner (inset) shadow using TextOptions.Shadow.PresetType = PresetShadowType.InsideCenter. The file is saved as XLSX.
// Keywords: Aspose.Cells | C# | .NET | inner shadow | PresetShadowType.InsideCenter | shape text styling | Excel shape shadow | TextOptions.Shadow | Excel formatting
// Common Searches: Aspose.Cells inner shadow shape text | PresetShadowType InsideCenter C# example | add inset shadow to Excel shape text | how to set text shadow in Aspose.Cells | C# Excel shape text styling with shadow
// Developer Intent: Apply a subtle inset shadow to the text inside a rectangle shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Enhance label readability on dashboard shapes by adding a soft inner shadow that creates depth without external shading. | Standardize text appearance across multiple report shapes, ensuring a consistent inset shadow for a professional look. | Automate the styling of shape titles in generated spreadsheets, applying the same inner‑shadow settings in bulk.
// AI Prompts: Show how to modify shadow distance, blur, and transparency to produce stronger or softer inset effects on shape text. | Provide an example using PresetShadowType.InnerShadow (if supported) with custom color and opacity settings. | Explain how to apply the same inset shadow configuration to every shape on a worksheet using a foreach loop.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;
using System.Drawing;

// Shows how to add a rectangle shape to an Excel workbook, set its text font, and apply a subtle inner (inset) shadow using TextOptions.Shadow.PresetType = PresetShadowType.InsideCenter. The file is saved as XLSX.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape that will contain the text
        // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
        Shape shape = sheet.Shapes.AddRectangle(2, 2, 2, 2, 200, 100);
        shape.Text = "Inset Shadow";

        // Obtain the FontSetting for the shape's text and then the TextOptions
        FontSetting fontSetting = shape.Characters(0, shape.Text.Length);
        TextOptions textOptions = fontSetting.TextOptions;

        // Set basic font properties
        textOptions.Name = "Calibri";
        textOptions.Size = 36;
        textOptions.IsBold = true;
        textOptions.Color = Color.DarkBlue;

        // Apply an inner (inset) shadow effect to the text
        textOptions.Shadow.PresetType = PresetShadowType.InsideCenter; // subtle inset shadow
        textOptions.Shadow.Transparency = 0.4; // make the shadow semi‑transparent
        textOptions.Shadow.Blur = 5;          // slight blur for softness
        textOptions.Shadow.Distance = 2;      // small distance to keep it subtle

        // Save the workbook to an XLSX file
        workbook.Save("InnerShadowShapeText.xlsx", SaveFormat.Xlsx);
    }
}
