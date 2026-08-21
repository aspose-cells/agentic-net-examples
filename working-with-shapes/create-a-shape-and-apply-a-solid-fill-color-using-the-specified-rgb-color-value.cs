// Title: Add a Rectangle Shape with Custom RGB Solid Fill in Aspose.Cells (C#)
// Description: Shows how to create a workbook, insert a rectangle shape, set its FillType to Solid, apply an RGB color (e.g., 135,222,255), optionally add text, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# shape fill | solid fill RGB | AddRectangle | Excel shape color | FillType.Solid | Aspose.Cells Drawing | custom shape color .NET
// Common Searches: Aspose.Cells set shape fill color C# | How to apply RGB solid fill to a shape in Aspose.Cells | Add rectangle shape to Excel with Aspose.Cells .NET | Change shape FillType to Solid Aspose.Cells | Create colored shape in workbook using Aspose.Cells
// Developer Intent: Insert a rectangle shape into a worksheet and color it with a specific RGB solid fill using Aspose.Cells for .NET.
// Use Cases: Design a branded header banner with a corporate RGB color. | Highlight a summary section by placing a colored shape behind cells. | Build a legend box that matches chart series colors.
// AI Prompts: Generate C# code that adds an ellipse shape and sets its solid fill to RGB (255,0,0) using Aspose.Cells. | Show how to modify the fill of an existing Aspose.Cells shape to a linear gradient in C#. | Provide a sample that creates multiple shapes, each with a different RGB solid fill, on the same worksheet.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeSolidFillDemo
{
    // Shows how to create a workbook, insert a rectangle shape, set its FillType to Solid, apply an RGB color (e.g., 135,222,255), optionally add text, and save the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Define the RGB color values you want to apply
            int red = 135;
            int green = 222;
            int blue = 255;
            Color fillColor = Color.FromArgb(red, green, blue);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
            Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

            // Set the fill type to solid and apply the RGB color
            shape.Fill.FillType = FillType.Solid;
            shape.Fill.SolidFill.Color = fillColor;

            // Optionally set a text to verify the shape
            shape.Text = "Solid Fill Shape";

            // Save the workbook to a file
            workbook.Save("ShapeSolidFillDemo.xlsx");
        }
    }
}
