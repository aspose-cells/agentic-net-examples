// Title: Set Shape Shadow Color with ARGB (255,128,0,0) and 30% Transparency in Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a rectangle shape, and configures its ShadowEffect. It assigns a CellsColor using ARGB values (255,128,0,0), sets the shadow opacity to 30 %, and adjusts angle, blur, distance, and size before saving the file as ShadowEffectArgbTransparency.xlsx.
// Keywords: Aspose.Cells | C# shape shadow | ARGB color | shadow transparency | ShadowEffect | CellsColor | Excel drawing API | global
// Common Searches: Aspose.Cells set shadow ARGB | C# shape shadow transparency 30 percent | how to change shadow color in Aspose.Cells | apply custom shadow effect to Excel shape .NET | Aspose.Cells shadow angle blur distance
// Developer Intent: Apply a specific ARGB color and 30 % opacity to a shape’s shadow using Aspose.Cells for .NET.
// Use Cases: Highlight a header box with a semi‑transparent red shadow in financial reports. | Standardize shadow styling for all diagram elements across multiple worksheets. | Create a template where shadow opacity emphasizes key visual cues.
// AI Prompts: Write C# code that sets a shape’s shadow to ARGB (255,128,0,0) with 30 % transparency using Aspose.Cells. | Show how to apply the same custom shadow (color, angle, blur, distance) to several shapes in a workbook. | Explain how to reuse a CellsColor instance for shadow effects on multiple shapes in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowExample
{
    // This example creates a workbook, adds a rectangle shape, and configures its ShadowEffect. It assigns a CellsColor using ARGB values (255,128,0,0), sets the shadow opacity to 30 %, and adjusts angle, blur, distance, and size before saving the file as ShadowEffectArgbTransparency.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to demonstrate the shadow effect
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

            // Get the shadow effect of the shape
            ShadowEffect shadow = shape.ShadowEffect;

            // Create a CellsColor instance and set its ARGB value to (255,128,0,0)
            CellsColor shadowColor = workbook.CreateCellsColor();
            shadowColor.Argb = Color.FromArgb(255, 128, 0, 0).ToArgb(); // Opaque semi‑red

            // Apply the color to the shadow effect
            shadow.Color = shadowColor;

            // Set the shadow transparency to 30% (0.3)
            shadow.Transparency = 0.3;

            // Optionally set other shadow properties for better visibility
            shadow.Angle = 135;
            shadow.Blur = 20;
            shadow.Distance = 10;
            shadow.Size = 1.0;

            // Save the workbook
            workbook.Save("ShadowEffectArgbTransparency.xlsx");

            Console.WriteLine("Workbook saved with shadow color (ARGB 255,128,0,0) and 30% transparency.");
        }
    }
}
