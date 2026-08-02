// Title: Aspose.Cells for .NET: Set a 45° Dark‑Gray Shadow on a Shape (C# Example)
// Description: This C# sample creates a workbook, inserts a rectangle shape, and uses the ShadowEffect API to apply a 45‑degree angle, 10‑point distance, and dark‑gray color before saving the file as ShadowEffectDemo.xlsx.
// Keywords: Aspose.Cells | C# shape shadow | ShadowEffect API | 45 degree shadow | dark gray shadow | Excel shape formatting | shadow distance
// Common Searches: Aspose.Cells set shape shadow angle C# | how to change shadow distance in Aspose.Cells | dark gray shadow for Excel shape using .NET | C# example for configuring shape shadow in Aspose.Cells | ShadowEffect properties Aspose.Cells tutorial
// Developer Intent: Apply a shadow with a 45° angle, 10‑point offset, and dark‑gray color to a shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Enhance visual hierarchy in generated reports by adding depth to call‑out shapes. | Create consistent branding across worksheets with predefined shadow styling. | Produce presentation‑style Excel sheets where shapes need a subtle, professional shadow.
// AI Prompts: Generate code to add blur radius and transparency to the shape shadow in Aspose.Cells. | Show how to apply the same 45° dark‑gray shadow to every shape on a worksheet programmatically. | Explain how to read, modify, or remove an existing shape's shadow after the workbook is opened.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// This C# sample creates a workbook, inserts a rectangle shape, and uses the ShadowEffect API to apply a 45‑degree angle, 10‑point distance, and dark‑gray color before saving the file as ShadowEffectDemo.xlsx.
class ConfigureShadowEffect
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to demonstrate the shadow effect
        Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

        // Access the shape's shadow effect
        ShadowEffect shadow = shape.ShadowEffect;

        // Configure shadow properties
        shadow.Angle = 45;      // 45 degree angle
        shadow.Distance = 10;   // 10 point distance

        // Set shadow color to dark gray
        CellsColor darkGray = workbook.CreateCellsColor();
        darkGray.Color = Color.DarkGray;
        shadow.Color = darkGray;

        // Save the workbook
        workbook.Save("ShadowEffectDemo.xlsx");
    }
}
