// Title: Link a rectangle shape to a multi‑line cell with wrapping using Aspose.Cells for .NET
// Description: Shows how to add a rectangle shape, bind it to a cell that contains line‑feed characters, enable text wrapping, and automatically resize the shape to fit the wrapped content with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | shape linked to cell | multiline cell text | text wrapping in shape | fit shape to text | SetLinkedCell method | UpdateSelectedValue | Shape.FitToTextSize | rectangle shape in Excel
// Common Searches: Aspose.Cells link shape to cell with newline | wrap text in linked shape Aspose.Cells | auto resize shape to cell text .NET | C# rectangle shape bound to cell | display multi‑line cell value in shape
// Developer Intent: Bind a rectangle shape to a cell that includes line breaks and have the shape render the text with proper wrapping.
// Use Cases: Dynamic dashboards where shape titles update automatically from cells containing bullet points or paragraphs. | Report generators that need wrapped notes inside shapes to stay synchronized with source cells. | Printable forms that show multi‑line instructions inside linked shapes for consistent layout.
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle shape linked to a cell containing newline characters and enables text wrapping. | Explain the role of the two Boolean parameters in SetLinkedCell and how they affect automatic updates of shape text. | Provide step‑by‑step instructions to fit a linked shape to its multi‑line content using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsShapeLinkExample
{
    // Shows how to add a rectangle shape, bind it to a cell that contains line‑feed characters, enable text wrapping, and automatically resize the shape to fit the wrapped content with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put multi‑line text into cell A1 (use line‑feed character for line break)
            worksheet.Cells["A1"].PutValue("First line\nSecond line\nThird line");

            // Add a rectangle shape that will display the linked text
            // Parameters: upper left row, upper left column, row offset, column offset, width, height
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 200, 100);

            // Link the shape to cell A1 (true, true = set as linked and update automatically)
            shape.SetLinkedCell("A1", true, true);

            // Ensure the shape reflects the current cell value
            shape.UpdateSelectedValue();

            // Enable text wrapping so line breaks are shown correctly
            shape.TextBody.TextAlignment.IsTextWrapped = true;

            // Adjust the shape size to fit the wrapped text
            shape.FitToTextSize();

            // Save the workbook
            workbook.Save("ShapeLinkedToMultilineCell.xlsx");
        }
    }
}
