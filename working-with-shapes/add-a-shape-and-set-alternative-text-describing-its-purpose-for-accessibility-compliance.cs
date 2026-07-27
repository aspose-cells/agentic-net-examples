// Title: Add a rectangle shape with alternative text in an Excel workbook using Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, inserts a rectangle shape on the first worksheet, assigns descriptive AlternativeText for screen‑reader accessibility, and saves the file as ShapeWithAlternativeText.xlsx.
// Keywords: Aspose.Cells C# shape | add rectangle Aspose.Cells | alternative text Excel shape | accessibility Aspose.Cells .NET | shape AlternativeText property | Excel workbook automation
// Common Searches: Aspose.Cells set shape alternative text C# | how to add rectangle shape in Excel with Aspose.Cells | accessibility alternative text for Excel shapes .NET | Aspose.Cells shape example C#
// Developer Intent: Insert a shape into a worksheet and provide alternative text for accessibility compliance.
// Use Cases: Design a sales dashboard where each visual element includes descriptive alternative text for screen readers. | Generate compliance‑ready reports that automatically label shapes with purpose‑specific text. | Automate workbook creation for corporate standards that require accessibility metadata on all graphics.
// AI Prompts: Generate C# code to add a circle shape and set its alternative text with Aspose.Cells. | Show how to loop through all worksheet shapes and assign alternative text based on shape type using Aspose.Cells for .NET. | Explain the steps to export an Excel file that preserves shape alternative text for accessibility tools.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeAlternativeTextDemo
{
    // Creates a new Workbook, inserts a rectangle shape on the first worksheet, assigns descriptive AlternativeText for screen‑reader accessibility, and saves the file as ShapeWithAlternativeText.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height, shape type (0 = rectangle)
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

            // Set alternative text for accessibility compliance
            shape.AlternativeText = "A rectangle shape representing the sales region overview";

            // Save the workbook (lifecycle: save)
            workbook.Save("ShapeWithAlternativeText.xlsx");
        }
    }
}
