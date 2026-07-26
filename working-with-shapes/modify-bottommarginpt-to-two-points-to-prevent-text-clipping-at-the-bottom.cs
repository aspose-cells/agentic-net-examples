// Title: Aspose.Cells .NET: Set Shape Text BottomMarginPt to 2 pts to Prevent Clipping
// Description: Demonstrates how to create a workbook, add a rectangle shape, assign text, and configure the shape's TextBody.TextAlignment.BottomMarginPt property to 2 points so the text is fully visible and not cut off at the bottom. The workbook is saved as ShapeWithBottomMargin.xlsx.
// Keywords: Aspose.Cells shape bottom margin | BottomMarginPt .NET | shape text clipping Aspose | adjust shape text margin C# | Aspose.Cells rectangle shape | prevent bottom text cut off
// Common Searches: how to set bottom margin of shape text in Aspose.Cells | BottomMarginPt property example C# | prevent text clipping in Aspose.Cells shapes | increase bottom margin of shape text Aspose
// Developer Intent: Apply a 2‑point bottom margin to a shape's text to avoid clipping.
// Use Cases: Add labeled shapes to a report and ensure the text does not get truncated at the bottom. | Generate dynamic worksheets where each shape’s BottomMarginPt is standardized for consistent rendering. | Create a template that automatically adjusts shape text margins after inserting variable‑length content.
// AI Prompts: Write C# code using Aspose.Cells to add a rectangle shape and set its TextBody.TextAlignment.BottomMarginPt to 2 points. | Explain the effect of BottomMarginPt on text layout inside shapes and how to use it to stop bottom clipping. | Provide a C# loop that iterates through all shapes on a worksheet and sets each shape's BottomMarginPt to 2 points.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to create a workbook, add a rectangle shape, assign text, and configure the shape's TextBody.TextAlignment.BottomMarginPt property to 2 points so the text is fully visible and not cut off at the bottom. The workbook is saved as ShapeWithBottomMargin.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, height, width
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 0);
        shape.Text = "Sample text with increased bottom margin";

        // Set the bottom margin of the shape's text to 2 points
        shape.TextBody.TextAlignment.BottomMarginPt = 2.0;

        // Save the workbook to a file
        workbook.Save("ShapeWithBottomMargin.xlsx");
    }
}
