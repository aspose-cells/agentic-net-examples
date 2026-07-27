// Title: Read Shape Text Margins (Top, Left, Bottom, Right) with Aspose.Cells for .NET
// Description: Shows how to access a shape's ShapeTextAlignment in Aspose.Cells, retrieve the TopMarginPt, LeftMarginPt, BottomMarginPt, and RightMarginPt values, and output them using C#.
// Keywords: Aspose.Cells | C# | ShapeTextAlignment | TopMarginPt | LeftMarginPt | BottomMarginPt | RightMarginPt | shape text margins | read shape margins | retrieve shape margins | Aspose.Cells shape margins
// Common Searches: Aspose.Cells get shape text margins C# | How to read TopMarginPt of a shape in Aspose.Cells | Retrieve LeftMarginPt from ShapeTextAlignment | Shape margin properties Aspose.Cells .NET | C# example reading shape margins Aspose.Cells
// Developer Intent: Obtain the current top, left, bottom, and right margin values of a shape's text body via Aspose.Cells for .NET.
// Use Cases: Validate shape text layout before exporting the workbook to PDF or other formats. | Adjust worksheet layout dynamically by reading existing shape margins and modifying surrounding content. | Log margin settings for debugging or reporting purposes. | Generate reports that adapt to the margin configuration of shapes.
// AI Prompts: Write C# code using Aspose.Cells that iterates over all shapes in a worksheet and prints their TopMarginPt, LeftMarginPt, BottomMarginPt, and RightMarginPt values. | Show how to compare a shape's margin values with target thresholds and update them if they differ. | Create a script that extracts shape margin data from a workbook and saves the results to a CSV file.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Shows how to access a shape's ShapeTextAlignment in Aspose.Cells, retrieve the TopMarginPt, LeftMarginPt, BottomMarginPt, and RightMarginPt values, and output them using C#.
class RetrieveShapeMargins
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = sheet.Shapes.AddRectangle(1, 1, 100, 80, 0, 0);
        shape.Text = "Sample text";

        // Access the text alignment object of the shape
        ShapeTextAlignment alignment = shape.TextBody.TextAlignment;

        // (Optional) Set margin values for demonstration
        alignment.TopMarginPt = 5.0;
        alignment.LeftMarginPt = 4.0;
        alignment.BottomMarginPt = 3.0;
        alignment.RightMarginPt = 2.0;

        // Retrieve current margin values
        double topMargin = alignment.TopMarginPt;
        double leftMargin = alignment.LeftMarginPt;
        double bottomMargin = alignment.BottomMarginPt;
        double rightMargin = alignment.RightMarginPt;

        // Output the margin values
        Console.WriteLine($"TopMarginPt: {topMargin}");
        Console.WriteLine($"LeftMarginPt: {leftMargin}");
        Console.WriteLine($"BottomMarginPt: {bottomMargin}");
        Console.WriteLine($"RightMarginPt: {rightMargin}");

        // Save the workbook (if needed)
        workbook.Save("ShapeMargins.xlsx");
    }
}
