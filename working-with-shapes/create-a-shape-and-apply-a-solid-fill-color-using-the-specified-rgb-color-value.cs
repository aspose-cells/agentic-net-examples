// Title: Aspose.Cells for .NET: Add a Rectangle Shape with an RGB Background (C#)
// Description: This C# sample creates a new workbook, inserts a rectangle on the first worksheet, sets its FillType to Solid, applies the RGB color (135,222,255) via Color.FromArgb, optionally adds text, and saves the result as ShapeSolidFillDemo.xlsx.
// Keywords: Aspose.Cells C# shape | rectangle shape | solid fill | RGB color | Color.FromArgb | FillType.Solid | add shape Aspose.Cells | shape formatting | Excel shape fill | programmatic shape color
// Common Searches: Aspose.Cells set shape color C# | how to add rectangle with custom fill in Aspose.Cells | apply RGB fill to shape in .NET Excel library | change shape fill type to solid using Aspose.Cells | C# code for shape solid fill Aspose.Cells
// Developer Intent: Insert a shape into a worksheet and color it with a specific RGB solid fill.
// Use Cases: Mark sections in an auto‑generated report with colored boxes. | Create a legend or key that uses a custom background hue. | Overlay a shape to emphasize a data range in an Excel file.
// AI Prompts: Write C# code that adds a circle shape with a solid fill of RGB(200,150,100) using Aspose.Cells. | Show how to modify the fill of an existing shape to a gradient based on two user‑provided colors in Aspose.Cells .NET. | Explain how to adjust the opacity of a shape’s solid fill programmatically with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# sample creates a new workbook, inserts a rectangle on the first worksheet, sets its FillType to Solid, applies the RGB color (135,222,255) via Color.FromArgb, optionally adds text, and saves the result as ShapeSolidFillDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, row offset, column offset, width, height
        Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

        // Set the fill type to solid
        shape.Fill.FillType = FillType.Solid;

        // Apply a solid fill color using an RGB value (e.g., 135, 222, 255)
        shape.Fill.SolidFill.Color = Color.FromArgb(135, 222, 255);

        // Optional: add some text to the shape
        shape.Text = "Solid Fill Shape";

        // Save the workbook to a file
        workbook.Save("ShapeSolidFillDemo.xlsx");
    }
}
