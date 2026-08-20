// Title: Add a Rectangle Shape with Alternative Text in Excel using Aspose.Cells for .NET
// Description: Shows how to create a workbook, insert a rectangle shape, assign alternative text for screen‑reader accessibility, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells shape | add rectangle Aspose.Cells | alternative text Excel shape .NET | Excel accessibility shapes | C# Aspose.Cells shape example | set AltText Aspose.Cells | worksheet.Shapes.AddRectangle | accessibility compliance Excel
// Common Searches: Aspose.Cells add rectangle shape C# | set alternative text for Excel shape .NET | Excel shape accessibility example | how to add alt text to shape using Aspose.Cells | C# code for shape alternative text in Excel
// Developer Intent: Insert a shape into an Excel worksheet and provide descriptive alternative text to meet accessibility standards.
// Use Cases: Highlight a sales region in a generated report with a rectangle that screen readers can describe. | Annotate charts programmatically with shapes that include alt text for compliance with WCAG. | Build interactive dashboards where each visual element conveys meaning through alternative text.
// AI Prompts: Generate C# code to add different shape types (ellipse, line, arrow) and set alternative text with Aspose.Cells. | Provide a snippet that reads and updates the AlternativeText of an existing shape in a workbook. | Explain how to assign alternative text to multiple shapes in a worksheet in a single pass using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeExample
{
    // Shows how to create a workbook, insert a rectangle shape, assign alternative text for screen‑reader accessibility, and save the file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet (row, column, upperLeftRow, upperLeftColumn, width, height, shapeIndex)
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

            // Set alternative text for accessibility
            shape.AlternativeText = "This rectangle highlights the sales region for Q1";

            // Save the workbook to a file
            workbook.Save("ShapeWithAltText.xlsx");
        }
    }
}
