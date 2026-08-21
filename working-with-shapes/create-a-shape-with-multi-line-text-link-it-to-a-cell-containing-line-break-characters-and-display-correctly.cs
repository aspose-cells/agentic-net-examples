// Title: Link a rectangle shape to a multi‑line cell with wrapped text – Aspose.Cells for .NET
// Description: Shows how to insert multi‑line text into a worksheet cell, add a rectangle shape, bind the shape to the cell, enable text wrapping so line breaks appear inside the shape, and save the workbook using Aspose.Cells for .NET (C#).
// Keywords: Aspose.Cells | C# | .NET | SetLinkedCell | UpdateSelectedValue | shape text wrapping | multiline cell | rectangle shape | linked shape | Excel shape API
// Common Searches: Aspose.Cells link shape to cell with line breaks | C# wrap text in linked shape Aspose.Cells | SetLinkedCell multiline text Aspose.Cells | Update shape after linking cell Aspose.Cells | Display multi‑line text in rectangle shape .NET
// Developer Intent: Bind a shape to a cell that contains line‑break characters and ensure the shape displays the text with proper wrapping.
// Use Cases: Create a report where a shape mirrors a cell's multi‑line description while preserving line breaks. | Build a dashboard that shows notes from cells inside shapes, updating automatically when the source cells change. | Design an Excel template where linked shapes act as captions that wrap text from source cells.
// AI Prompts: Write C# code with Aspose.Cells to add a rectangle shape, link it to cell A1 containing '\n' line breaks, enable text wrapping, and save the workbook. | Explain how SetLinkedCell and UpdateSelectedValue work together to reflect multi‑line cell content in a linked shape using Aspose.Cells for .NET. | Provide troubleshooting steps if a shape linked to a multi‑line cell does not show line breaks after linking.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsShapeLinkExample
{
    // Shows how to insert multi‑line text into a worksheet cell, add a rectangle shape, bind the shape to the cell, enable text wrapping so line breaks appear inside the shape, and save the workbook using Aspose.Cells for .NET (C#).
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put multi‑line text into cell A1 (use \n for line breaks)
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("First line\nSecond line\nThird line");

            // Add a rectangle shape that will display the cell's text
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 200, 100);

            // Link the shape to the cell containing the multi‑line text
            shape.SetLinkedCell("A1", true, true);
            // Refresh the shape's displayed value from the linked cell
            shape.UpdateSelectedValue();

            // Enable text wrapping so line breaks are shown correctly inside the shape
            shape.TextBody.TextAlignment.IsTextWrapped = true;

            // Save the workbook
            workbook.Save("ShapeLinkedToMultilineCell.xlsx");
        }
    }
}
