// Title: Apply a 45° Dark Gray Shadow (10 pt distance) to a Rectangle Shape in Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a rectangle shape, and configure its ShadowEffect—setting the angle to 45°, distance to 10 points, and color to dark gray—using Aspose.Cells for .NET, then save the Excel file.
// Keywords: Aspose.Cells | C# | .NET | Excel shape shadow | ShadowEffect | rectangle shape | shadow angle | shadow distance | shadow color | dark gray shadow | 45 degree shadow | programmatic Excel styling | Aspose.Cells API
// Common Searches: Aspose.Cells set shape shadow angle | C# Aspose.Cells shadow distance | change shape shadow color Aspose.Cells | apply shadow to rectangle in Excel using Aspose.Cells | Aspose.Cells shadow effect example
// Developer Intent: Create a rectangle shape in a worksheet and apply a shadow with a 45° angle, 10‑point offset, and dark gray color via the Aspose.Cells API.
// Use Cases: Enhance visual hierarchy in automated Excel reports by adding depth to annotation shapes. | Programmatically style diagram elements to match corporate branding in exported workbooks. | Highlight key data points in dashboards by applying consistent shadow effects to shapes.
// AI Prompts: Generate C# code with Aspose.Cells that adds a circle shape and applies a 30° angle, 5‑point distance, light‑blue shadow. | Explain how to adjust shadow blur and transparency for shapes using Aspose.Cells for .NET. | Provide a step‑by‑step guide to assign different shadow colors to multiple shapes on the same worksheet.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add a rectangle shape, and configure its ShadowEffect—setting the angle to 45°, distance to 10 points, and color to dark gray—using Aspose.Cells for .NET, then save the Excel file.
class ConfigureShadowEffect
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

        // Access the shape's shadow effect
        ShadowEffect shadow = shape.ShadowEffect;

        // Set the required shadow properties
        shadow.Angle = 45;               // 45 degree angle
        shadow.Distance = 10;            // 10 points distance

        // Create a CellsColor for dark gray and assign it to the shadow
        CellsColor darkGray = workbook.CreateCellsColor();
        darkGray.Color = Color.DarkGray; // System.Drawing.Color.DarkGray
        shadow.Color = darkGray;

        // Save the workbook to a file
        workbook.Save("ShadowEffectDemo.xlsx");
    }
}
