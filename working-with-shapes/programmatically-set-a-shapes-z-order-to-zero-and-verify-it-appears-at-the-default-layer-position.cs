// Title: Aspose.Cells for .NET: Set Shape Z‑Order to Default Layer (0) and Verify
// Description: Creates a workbook, adds a rectangle shape, sets its ZOrderPosition property to 0 to place it on the default layer, prints the resulting Z‑order value, and saves the file as an XLSX document.
// Keywords: Aspose.Cells ZOrderPosition | C# shape Z-order | default layer shape Aspose.Cells | reset shape Z-order .NET | Excel shape layering | Aspose.Cells shape order | C# workbook shape Z-order
// Common Searches: How to set a shape's ZOrderPosition to 0 in Aspose.Cells | Get shape Z-order value after assigning it in C# | Move a shape to the back of the drawing stack programmatically with Aspose.Cells | Check default layer of a shape in an Excel workbook using Aspose.Cells | Aspose.Cells shape layering examples
// Developer Intent: Programmatically set a worksheet shape's ZOrderPosition to 0 (default layer) and confirm the value.
// Use Cases: Place newly added shapes behind existing content to avoid visual overlap. | Validate shape layering before exporting reports to ensure consistent appearance across Excel viewers. | Normalize Z‑order across multiple worksheets during automated report generation.
// AI Prompts: Show C# code that sets a shape's ZOrderPosition to 0 and confirms it is on the default layer with Aspose.Cells. | Provide an example that adds several shapes, assigns custom ZOrderPosition values, and prints their order to verify layering in Aspose.Cells for .NET. | Explain how ZOrderPosition interacts with other shapes and how to reset a shape to the default layer in an Excel workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // Creates a workbook, adds a rectangle shape, sets its ZOrderPosition property to 0 to place it on the default layer, prints the resulting Z‑order value, and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddRectangle(5, 5, 0, 0, 100, 100);

            // Set the shape's Z-order position to zero (default layer)
            shape.ZOrderPosition = 0;

            // Verify the Z-order position
            Console.WriteLine("Shape ZOrderPosition after setting to zero: " + shape.ZOrderPosition);

            // Save the workbook to a file
            workbook.Save("ShapeZOrderZeroDemo.xlsx");
        }
    }
}
