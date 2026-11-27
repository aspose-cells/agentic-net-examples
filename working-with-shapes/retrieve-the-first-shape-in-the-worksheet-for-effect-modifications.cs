// Title: How to retrieve the first shape in a worksheet and modify its ShadowEffect and GlowEffect with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get first shape from worksheet and change shadow properties | set glow effect on a worksheet shape using Aspose.Cells .NET | add rectangle shape when worksheet has no shapes Aspose.Cells example | modify shape shadow angle blur distance transparency Aspose.Cells C# | retrieve shape at index 0 and apply visual effects Aspose.Cells
// Tags: retrieve first worksheet shape Aspose.Cells | modify shape shadow effect Aspose.Cells .NET | apply glow effect to shape Aspose.Cells C# | add rectangle shape if none Aspose.Cells | shape effect customization Aspose.Cells workbook

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace FirstShapeEffectDemo
{
    // The example creates a new workbook, ensures a rectangle shape exists on the first worksheet, retrieves the first shape, and then customizes its ShadowEffect (angle, blur, distance, transparency, size) and GlowEffect (color, size, transparency) before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape.
            // Add a rectangle shape if the worksheet has no shapes.
            if (worksheet.Shapes.Count == 0)
            {
                // Parameters: upper left row, upper left column, top, left, height, width
                // Using row=1, column=0, top=0, left=0, height=100, width=200
                worksheet.Shapes.AddRectangle(1, 0, 0, 0, 100, 200);
                Console.WriteLine("Added a rectangle shape to the worksheet.");
            }

            // Retrieve the first shape
            Shape firstShape = worksheet.Shapes[0];
            Console.WriteLine($"Retrieved first shape: Type={firstShape.Type}, Name={firstShape.Name}");

            // Modify effect properties - ShadowEffect as an example
            ShadowEffect shadow = firstShape.ShadowEffect;

            // Set various shadow properties
            shadow.Angle = 135;          // Lighting angle in degrees
            shadow.Blur = 20;            // Blur amount (0-100 points)
            shadow.Distance = 15;        // Distance from shape (0-200 points)
            shadow.Transparency = 0.4;   // 40% transparent
            shadow.Size = 1.2;           // Size factor (0-2.0)

            // Optionally modify glow effect as well
            GlowEffect glow = firstShape.Glow;
            CellsColor glowColor = workbook.CreateCellsColor();
            glowColor.Color = System.Drawing.Color.Purple;
            glow.Color = glowColor;
            glow.Size = 8;               // Glow radius in points
            glow.Transparency = 0.3;     // 30% transparent

            Console.WriteLine("Modified ShadowEffect and GlowEffect of the first shape.");

            // Save the workbook to verify changes
            string outputPath = "FirstShapeEffectDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
