// Title: C# – Apply ARGB (255,128,0,0) Shadow Color and 30% Transparency to a Shape with Aspose.Cells
// Description: This example creates an Excel workbook, inserts a rectangle, accesses its ShadowEffect, assigns a dark‑red ARGB color (255,128,0,0), sets the opacity to 30 %, and optionally tweaks angle, blur and distance before saving as ShadowColorTransparency.xlsx.
// Keywords: Aspose.Cells C# shape shadow | ARGB shadow color Aspose.Cells | shadow transparency .NET | Excel shape shadow properties | Aspose.Cells shadow angle blur distance | C# Excel graphics styling
// Common Searches: how to change shape shadow color in Aspose.Cells | set shadow opacity to 30 percent using Aspose.Cells .NET | custom ARGB values for Excel shape shadows | configure shadow angle and blur for a rectangle in Aspose.Cells
// Developer Intent: Add a custom ARGB hue and 30 % opacity to a shape's shadow in an Excel file via Aspose.Cells for .NET.
// Use Cases: Highlight key cells with a dark‑red, semi‑transparent shadow for visual emphasis. | Create 3‑D‑like annotations on charts by adjusting shadow angle, blur, and distance. | Apply a consistent shadow style across all shapes in a corporate report template.
// AI Prompts: Generate C# code that sets a shape's shadow to ARGB (255,128,0,0) with 30% transparency using Aspose.Cells. | Show how to modify shadow angle, blur radius, and offset after defining color and opacity for a rectangle. | Explain reusing a CellsColor instance for multiple shape shadows while varying transparency levels.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// This example creates an Excel workbook, inserts a rectangle, accesses its ShadowEffect, assigns a dark‑red ARGB color (255,128,0,0), sets the opacity to 30 %, and optionally tweaks angle, blur and distance before saving as ShadowColorTransparency.xlsx.
class SetShadowExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to demonstrate the shadow effect
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

        // Retrieve the shadow effect of the shape
        ShadowEffect shadow = shape.ShadowEffect;

        // Create a CellsColor instance and set its ARGB value to (255,128,0,0)
        CellsColor shadowColor = workbook.CreateCellsColor();
        shadowColor.Argb = Color.FromArgb(255, 128, 0, 0).ToArgb(); // Opaque dark red

        // Apply the color and set transparency to 30%
        shadow.Color = shadowColor;
        shadow.Transparency = 0.3; // 30% transparent

        // Optional: configure additional shadow properties
        shadow.Angle = 135;
        shadow.Blur = 20;
        shadow.Distance = 10;

        // Save the workbook
        workbook.Save("ShadowColorTransparency.xlsx");
    }
}
