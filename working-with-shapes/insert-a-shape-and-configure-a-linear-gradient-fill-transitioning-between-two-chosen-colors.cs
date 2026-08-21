// Title: Aspose.Cells for .NET – Add a rectangle shape with a linear two‑color gradient fill (C#)
// Description: Creates a new workbook, inserts a rectangle shape on the first worksheet, sets the shape's fill to a linear gradient at 45°, applies a horizontal two‑color gradient from LightSkyBlue to DarkBlue, and saves the file as an XLSX document.
// Keywords: Aspose.Cells C# shape gradient | linear gradient fill Aspose.Cells | two‑color gradient rectangle .NET | GradientFillType.Linear | GradientStyleType.Horizontal | shape fill type gradient | gradient angle Aspose.Cells | add rectangle shape C# | Aspose.Cells gradient example
// Common Searches: how to add a rectangle shape with linear gradient in Aspose.Cells C# | set two‑color gradient fill for a shape using Aspose.Cells .NET | Aspose.Cells linear gradient angle 45 degrees | apply horizontal gradient to a shape in a workbook | C# code for gradient fill on Aspose.Cells shape
// Developer Intent: Insert a shape into a worksheet and apply a linear two‑color gradient fill.
// Use Cases: Create a branded header banner with a LightSkyBlue‑to‑DarkBlue gradient. | Highlight a chart area by overlaying a gradient‑filled rectangle. | Generate report templates where all shapes share a consistent gradient style.
// AI Prompts: Generate C# code that adds an ellipse shape with a vertical linear gradient from red to orange using Aspose.Cells. | Show how to change the gradient angle and style of an existing shape's fill in an Aspose.Cells workbook. | Provide an example of applying a three‑color gradient to a shape and saving the workbook with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGradientShapeDemo
{
    // Creates a new workbook, inserts a rectangle shape on the first worksheet, sets the shape's fill to a linear gradient at 45°, applies a horizontal two‑color gradient from LightSkyBlue to DarkBlue, and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape (you can adjust position and size as needed)
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape shape = worksheet.Shapes.AddRectangle(2, 1, 50, 50, 200, 100);

            // Set the fill type of the shape to Gradient to enable gradient properties
            shape.Fill.FillType = FillType.Gradient;

            // Obtain the GradientFill object from the shape's fill
            GradientFill gradientFill = shape.Fill.GradientFill;

            // Define a linear gradient with a specific angle (e.g., 45 degrees)
            // GradientDirectionType is ignored for linear gradients, but a value must be supplied
            gradientFill.SetGradient(GradientFillType.Linear, 45.0, GradientDirectionType.FromCenter);

            // Apply a two‑color gradient (e.g., from LightSkyBlue to DarkBlue)
            // GradientStyleType.Horizontal defines the shading direction for the linear gradient
            gradientFill.SetTwoColorGradient(
                Color.LightSkyBlue,   // First color
                Color.DarkBlue,       // Second color
                GradientStyleType.Horizontal,
                1);                   // Variant (1‑4, choose 1 for default)

            // Save the workbook to a file
            workbook.Save("ShapeWithLinearGradient.xlsx");

            Console.WriteLine("Workbook saved with a shape that has a linear gradient fill.");
        }
    }
}
