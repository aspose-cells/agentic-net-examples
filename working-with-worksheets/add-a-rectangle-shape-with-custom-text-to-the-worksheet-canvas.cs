// Title: Insert a Rectangle Shape with Text in Excel via Aspose.Cells for .NET
// Description: C# sample that creates a workbook, adds a rectangle to the first worksheet at row 2, column 2, assigns custom text, applies a light‑green solid fill and a solid border, then saves the file as RectangleWithText.xlsx.
// Keywords: Aspose.Cells rectangle shape C# | Excel shape text Aspose | add rectangle to worksheet | custom shape styling Aspose.Cells | Aspose.Cells Drawing API | set shape fill color .NET | shape line weight Aspose
// Common Searches: how to add a rectangle with text using Aspose.Cells | Aspose.Cells C# shape fill color example | set border style for rectangle shape in Excel with Aspose | draw shapes on worksheet canvas Aspose.Cells | C# code to place rectangle at specific cells Aspose
// Developer Intent: Place a rectangle on a worksheet, embed custom label, and format its appearance programmatically.
// Use Cases: Create visual section dividers in financial reports. | Design button‑like areas on a dashboard that guide users to other data views. | Build printable forms with labeled boxes for user input.
// AI Prompts: Generate C# code that adds a rectangle shape with wrapped multiline text and centers it in a cell using Aspose.Cells. | Show how to create several rectangles with different colors and export the workbook to PDF. | Explain how to assign a hyperlink to a shape and handle click actions in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsRectangleWithText
{
    // C# sample that creates a workbook, adds a rectangle to the first worksheet at row 2, column 2, assigns custom text, applies a light‑green solid fill and a solid border, then saves the file as RectangleWithText.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet canvas
            // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height (pixels), width (pixels)
            RectangleShape rectangle = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 120, 200);

            // Set custom text for the rectangle
            rectangle.Text = "Custom rectangle text";

            // Optionally, customize appearance (fill color, line style, etc.)
            rectangle.Fill.FillType = FillType.Solid;
            rectangle.Fill.SolidFill.Color = System.Drawing.Color.LightGreen;
            rectangle.Line.DashStyle = MsoLineDashStyle.Solid;
            rectangle.Line.Weight = 1.5;

            // Save the workbook to a file
            workbook.Save("RectangleWithText.xlsx");
        }
    }
}
