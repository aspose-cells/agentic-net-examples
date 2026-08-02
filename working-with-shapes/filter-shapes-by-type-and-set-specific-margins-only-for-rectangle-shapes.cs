// Title: Aspose.Cells for .NET – Set Custom Text Margins on Rectangle Shapes in Excel
// Description: C# example that loads an Excel workbook, iterates through each worksheet, filters shapes to rectangles, sets top, bottom, left, and right text margins (points) via ShapeTextAlignment, disables automatic margin calculation, and saves the updated file.
// Keywords: Aspose.Cells | C# shape margins | Excel rectangle shape | ShapeTextAlignment | set text padding | filter shapes by type | AutoShapeType.Rectangle | disable auto margin | Aspose.Cells .NET | Excel shape formatting
// Common Searches: Aspose.Cells set rectangle shape margins | C# change text padding of Excel shapes | filter shapes by type Aspose.Cells | disable auto margin Aspose.Cells shape | how to set shape text margins in .NET
// Developer Intent: Apply specific top, bottom, left, and right text margins only to rectangle shapes in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Standardize padding of rectangle callout shapes in financial report templates. | Prepare Excel workbooks with uniform shape margins before converting to PDF. | Automate layout adjustments for dashboard widgets stored as rectangle shapes. | Ensure printable Excel charts have consistent text spacing inside rectangle containers.
// AI Prompts: Write C# code using Aspose.Cells to set 12‑pt margins on all ellipse shapes in a workbook. | Show how to list each shape's name and type while applying custom margins with Aspose.Cells. | Create a script that reads margin values from a JSON file and applies them to rectangle shapes. | Provide a PowerShell example that calls Aspose.Cells to adjust text margins for rectangle shapes.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsShapeMarginExample
{
    // C# example that loads an Excel workbook, iterates through each worksheet, filters shapes to rectangles, sets top, bottom, left, and right text margins (points) via ShapeTextAlignment, disables automatic margin calculation, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Check if the shape is a rectangle (either by type or by AutoShapeType)
                    bool isRectangle = shape is RectangleShape ||
                                       shape.Type == AutoShapeType.Rectangle;

                    if (isRectangle)
                    {
                        // Access the text alignment object of the shape
                        ShapeTextAlignment alignment = shape.TextBody.TextAlignment;

                        // Set specific margins (values are in points)
                        alignment.TopMarginPt = 5.0;      // Top margin
                        alignment.BottomMarginPt = 5.0;   // Bottom margin
                        alignment.LeftMarginPt = 8.0;     // Left margin
                        alignment.RightMarginPt = 8.0;    // Right margin

                        // Optionally disable automatic margin calculation
                        alignment.IsAutoMargin = false;
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
