// Title: C# – Apply fixed text margins to rectangle shapes in an Excel file with Aspose.Cells
// Description: The sample opens an Excel workbook, walks through each worksheet and its shapes, selects only RectangleShape objects, sets explicit top, bottom, left, and right margins via ShapeTextAlignment, disables automatic margin calculation, and saves the modified file.
// Keywords: Aspose.Cells | C# shape margins | RectangleShape | ShapeTextAlignment | Excel shape padding | disable auto margin | set text margins .NET | iterate worksheet shapes | custom shape formatting
// Common Searches: How to change padding inside rectangle shapes in Excel using Aspose.Cells C# | Aspose.Cells iterate shapes and set fixed margins | Disable automatic margins for specific shape types with Aspose.Cells | Set top and left margin for rectangle shape text in .NET | Batch update shape text alignment across worksheets Aspose.Cells
// Developer Intent: Programmatically adjust the text padding of rectangle shapes in an Excel workbook with Aspose.Cells for .NET.
// Use Cases: Prepare a corporate template where all rectangle callout boxes must have uniform 5‑pt padding. | Generate printable reports that require consistent text positioning inside rectangle shapes. | Automate the cleanup of legacy workbooks to enforce branding‑compliant shape margins.
// AI Prompts: Write C# code that sets 8‑point margins for all oval shapes and keeps auto‑margin enabled using Aspose.Cells. | Show how to log the name and original margin values of each rectangle shape before updating them. | Create a script that toggles IsAutoMargin on for rectangle shapes after custom margins have been applied.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// The sample opens an Excel workbook, walks through each worksheet and its shapes, selects only RectangleShape objects, sets explicit top, bottom, left, and right margins via ShapeTextAlignment, disables automatic margin calculation, and saves the modified file.
class ShapeMarginProcessor
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Iterate through all shapes on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Process only rectangle shapes
                if (shape is RectangleShape rectangle)
                {
                    // Access the text alignment object of the rectangle
                    ShapeTextAlignment alignment = rectangle.TextBody.TextAlignment;

                    // Set explicit margins (in points) and disable auto‑margin
                    alignment.TopMarginPt = 5.0;
                    alignment.BottomMarginPt = 5.0;
                    alignment.LeftMarginPt = 5.0;
                    alignment.RightMarginPt = 5.0;
                    alignment.IsAutoMargin = false;
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
