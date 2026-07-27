// Title: Add a Rectangle Shape with an External Hyperlink using Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, inserts a rectangle shape at cell B2 (row 1, column 1) sized 100 × 100 points, attaches the URL https://www.example.com/ to the shape, and saves the file as ShapeWithHyperlink.xlsx.
// Keywords: Aspose.Cells C# shape hyperlink | add rectangle shape Excel .NET | shape AddHyperlink example | clickable shape Excel workbook | Aspose.Cells external link shape
// Common Searches: Aspose.Cells add hyperlink to shape C# | how to insert a clickable rectangle in Excel using Aspose.Cells | C# code for shape with URL in Excel file | Aspose.Cells shape AddHyperlink method usage | save workbook with linked shape Aspose.Cells
// Developer Intent: Insert a shape into a worksheet and make it open a specified website when the user clicks it.
// Use Cases: Dashboard button that launches an online help page. | Company logo that redirects to the corporate website. | Report element linking to detailed documentation or a web portal.
// AI Prompts: Generate C# code that adds multiple shapes, each with a different external URL, using Aspose.Cells. | Show how to apply fill color, border style, and text to a rectangle shape while keeping its hyperlink functional. | Explain how to attach a hyperlink to a circular shape that points to a PDF hosted online, using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsHyperlinkExample
{
    // Creates a new Workbook, inserts a rectangle shape at cell B2 (row 1, column 1) sized 100 × 100 points, attaches the URL https://www.example.com/ to the shape, and saves the file as ShapeWithHyperlink.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet (row, column, upper left X, upper left Y, width, height)
            // Here we place it at row 1, column 1 with size 100x100 points
            Shape shape = worksheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);

            // Attach a hyperlink to the shape that points to an external website
            shape.AddHyperlink("https://www.example.com/");

            // Save the workbook to a file
            workbook.Save("ShapeWithHyperlink.xlsx");
        }
    }
}
