// Title: Aspose.Cells C# – Apply an Outer Shadow Preset with 40% Transparency to a Shape
// Description: Create a workbook, add a rectangle shape, set its ShadowEffect to the OffsetBottom preset, adjust transparency to 0.4 (40 %), and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells shape shadow | C# outer shadow preset | shadow transparency Aspose.Cells | PresetShadowType OffsetBottom | Aspose.Cells .NET shape formatting | Excel shape shadow effect
// Common Searches: how to add outer shadow to a shape with Aspose.Cells C# | set shadow transparency for Excel shape using Aspose.Cells | Aspose.Cells preset shadow types example | C# code for shape shadow effect in Aspose.Cells | apply 40% transparent shadow to rectangle in Aspose.Cells
// Developer Intent: Add an outer shadow preset to a worksheet shape and set its transparency to 40 % with Aspose.Cells for .NET.
// Use Cases: Highlight key cells by adding a subtle bottom‑offset shadow to button‑style rectangles. | Standardize visual style across generated reports by applying the same semi‑transparent outer shadow to all diagram shapes. | Create marketing worksheets where shapes have a light shadow to improve depth perception and readability.
// AI Prompts: Generate C# code that applies different outer shadow presets to a shape and sets each transparency level using Aspose.Cells. | Show how to retrieve and modify the ShadowEffect of existing shapes in an Excel file with Aspose.Cells. | Explain how to loop through all shapes on a worksheet and assign the OffsetBottom preset with 40% transparency.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowDemo
{
    // Create a workbook, add a rectangle shape, set its ShadowEffect to the OffsetBottom preset, adjust transparency to 0.4 (40 %), and save the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

            // Get the shadow effect of the shape
            ShadowEffect shadow = shape.ShadowEffect;

            // Apply an outer shadow preset (e.g., OffsetBottom)
            shadow.PresetType = PresetShadowType.OffsetBottom;

            // Set the transparency of the shadow to 40% (0.4)
            shadow.Transparency = 0.4;

            // Save the workbook (lifecycle save)
            workbook.Save("ShapeWithOuterShadow.xlsx");

            Console.WriteLine("Workbook saved with outer shadow preset and 40% transparency.");
        }
    }
}
