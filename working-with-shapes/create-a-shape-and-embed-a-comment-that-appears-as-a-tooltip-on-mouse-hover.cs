// Title: Add a Rectangle Shape with Hover Tooltip and Cell Comment using Aspose.Cells for .NET (C#)
// Description: This example shows how to create a new workbook, insert a rectangle shape, set its fill and line colors, assign hover text via the shape's AlternativeText property, add a cell comment that appears as a tooltip, and save the file as an XLSX workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Excel shape tooltip | AlternativeText property | cell comment hover | rectangle shape | sample code | API example | Excel automation
// Common Searches: Aspose.Cells add tooltip to shape C# | How to set AlternativeText for a shape in Aspose.Cells | Create cell comment that shows on hover with Aspose.Cells | Rectangle shape with hover text in Excel using Aspose.Cells | Aspose.Cells shape tooltip example
// Developer Intent: Insert a rectangle shape and a cell comment that display tooltip text when the user hovers over them in an Excel workbook.
// Use Cases: Provide explanatory notes for diagram elements without cluttering the sheet. | Add guidance for form fields that appears on mouse‑over. | Create interactive documentation inside a spreadsheet by embedding hover text in shapes and cells.
// AI Prompts: Generate C# code with Aspose.Cells to add a rectangle shape and set its AlternativeText for a hover tooltip. | Explain how to customize the appearance of a shape tooltip and a cell comment in an Aspose.Cells workbook. | Show how to read, modify, or delete the tooltip text of an existing shape using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentShapeDemo
{
    // This example shows how to create a new workbook, insert a rectangle shape, set its fill and line colors, assign hover text via the shape's AlternativeText property, add a cell comment that appears as a tooltip, and save the file as an XLSX workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Add a comment to cell B2 (appears as a tooltip)
            // -------------------------------------------------
            int commentIndex = worksheet.Comments.Add("B2");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "This is a tooltip comment for cell B2";

            // -------------------------------------------------
            // 2. Add a rectangle shape to the worksheet
            // -------------------------------------------------
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            Shape rectangle = worksheet.Shapes.AddRectangle(5, 2, 0, 0, 150, 80);
            rectangle.Name = "MyRectangle";

            // Set visual properties (optional)
            rectangle.FillFormat.ForeColor = System.Drawing.Color.LightGreen;
            rectangle.LineFormat.ForeColor = System.Drawing.Color.DarkGreen;

            // -------------------------------------------------
            // 3. Embed a tooltip into the shape using AlternativeText
            // -------------------------------------------------
            rectangle.AlternativeText = "This rectangle shows a tooltip when hovered";

            // -------------------------------------------------
            // 4. Save the workbook to an XLSX file
            // -------------------------------------------------
            workbook.Save("ShapeWithTooltip.xlsx");
        }
    }
}
