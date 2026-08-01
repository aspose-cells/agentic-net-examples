// Title: Aspose.Cells for .NET: Apply Solid Color and Transparency to a TextBox Shape
// Description: Shows how to create a workbook, insert a TextBox shape, set its FillType to Solid, assign an ARGB color, adjust the SolidFill.Transparency (0‑1 range), and save the result as an XLSX file.
// Keywords: Aspose.Cells | C# textbox shape fill | solid fill transparency | ARGB color Excel | shape background color .NET | Excel textbox styling | FillType.Solid | SolidFill.Transparency
// Common Searches: set textbox background color Aspose.Cells C# | how to make textbox shape transparent in Excel using Aspose | solid fill for shapes Aspose.Cells .NET | change shape fill opacity programmatically | apply ARGB color to Excel shape with Aspose
// Developer Intent: Set a TextBox shape’s background to a solid ARGB color and optionally control its opacity in an Excel workbook via Aspose.Cells for .NET.
// Use Cases: Design report headers with colored, semi‑transparent text boxes | Highlight data regions by overlaying a translucent textbox | Generate Excel templates where background colors are defined by ARGB values | Create dashboards with visual emphasis using partially opaque shapes
// AI Prompts: Write C# code using Aspose.Cells to add a TextBox shape and set its background to a solid color with 30% transparency. | Explain how to use Fill.FillType and SolidFill.Transparency to style a textbox in an Excel file. | Show an example of assigning an ARGB value to a shape’s fill and adjusting opacity in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, insert a TextBox shape, set its FillType to Solid, assign an ARGB color, adjust the SolidFill.Transparency (0‑1 range), and save the result as an XLSX file.
class SetTextboxBackground
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a textbox shape to the worksheet
        // Parameters: drawing type, upper left row, upper left column, top, left, width, height
        Shape textbox = sheet.Shapes.AddShape(MsoDrawingType.TextBox, 1, 0, 1, 0, 200, 100);
        textbox.Text = "Sample Text";

        // Configure the fill to be solid
        textbox.Fill.FillType = FillType.Solid;

        // Set the solid fill color (opaque part)
        textbox.Fill.SolidFill.Color = Color.FromArgb(255, 100, 150, 200); // ARGB where A=255 (opaque)

        // Set optional transparency (0.0 = fully opaque, 1.0 = fully transparent)
        textbox.Fill.SolidFill.Transparency = 0.3; // 30% transparent

        // Save the workbook
        workbook.Save("TextboxBackgroundSolidFill.xlsx");
    }
}
