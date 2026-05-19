using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowOffsetDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: add some shapes to the first worksheet for demonstration
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Shapes.AddRectangle(2, 0, 2, 0, 150, 100);
            sheet.Shapes.AddOval(5, 0, 5, 0, 120, 80);
            sheet.Shapes.AddTextBox(8, 0, 8, 0, 200, 60);

            // Define the uniform shadow offset distance (in points)
            double uniformDistance = 10.0;

            // Iterate through all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Iterate through all shapes in the current worksheet
                foreach (Shape shape in ws.Shapes)
                {
                    // Access the shape's ShadowEffect object
                    ShadowEffect shadow = shape.ShadowEffect;

                    // Apply a preset shadow type that uses offset (e.g., OffsetBottom)
                    shadow.PresetType = PresetShadowType.OffsetBottom;

                    // Set the uniform distance for the shadow offset
                    shadow.Distance = uniformDistance;

                    // Optional: you can also set other shadow properties for consistency
                    // shadow.Angle = 90;          // direction of the shadow
                    // shadow.Blur = 5;            // blur amount
                    // shadow.Transparency = 0.3; // transparency level
                }
            }

            // Save the workbook with the applied shadow settings
            workbook.Save("UniformShadowOffsetDemo.xlsx");
        }
    }
}