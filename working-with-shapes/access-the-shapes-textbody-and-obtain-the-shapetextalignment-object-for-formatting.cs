// Title: Retrieve and format a shape's TextBody alignment using Aspose.Cells for .NET
// Description: Demonstrates how to add a rectangle shape to a worksheet, access its TextBody, obtain the ShapeTextAlignment object, enable text wrapping, set custom margins, and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells shape TextBody | ShapeTextAlignment C# | text wrapping Aspose.Cells | shape margin settings .NET | format shape text Aspose.Cells
// Common Searches: how to get ShapeTextAlignment from a shape in Aspose.Cells | enable text wrap for rectangle shape Aspose.Cells .NET | set top bottom left right margins for shape text | Aspose.Cells shape text alignment example
// Developer Intent: Access a shape's TextBody, retrieve the ShapeTextAlignment object, and apply formatting such as wrapping and margin adjustments.
// Use Cases: Prepare reports where shape captions need consistent padding and line wrapping. | Create reusable templates with predefined text layout for charts, diagrams, or callouts. | Programmatically adjust shape text appearance in bulk before exporting to Excel.
// AI Prompts: Provide C# code that retrieves a shape's ShapeTextAlignment via TextBody and sets wrap and margin properties using Aspose.Cells. | Write a method that accepts a Shape object and configures its text alignment, wrap flag, and margin values. | Explain the steps to modify ShapeTextAlignment properties of a shape in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to add a rectangle shape to a worksheet, access its TextBody, obtain the ShapeTextAlignment object, enable text wrapping, set custom margins, and save the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 50);
        shape.Text = "Sample Text";

        // Access the ShapeTextAlignment object via the shape's TextBody
        ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;

        // Example formatting: enable text wrapping and set custom margins
        textAlignment.IsTextWrapped = true;
        textAlignment.TopMarginPt = 10;
        textAlignment.BottomMarginPt = 10;
        textAlignment.LeftMarginPt = 5;
        textAlignment.RightMarginPt = 5;

        // Save the workbook with the formatted shape
        workbook.Save("ShapeTextAlignmentDemo.xlsx");
    }
}
