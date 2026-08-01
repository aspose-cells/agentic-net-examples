// Title: Read and Preserve TextParagraph SpaceBefore (Points) in a TextBox Shape – Aspose.Cells for .NET
// Description: Demonstrates how to read the current SpaceBefore value of a TextParagraph in a TextBox shape, modify it using LineSpaceSizeType.Points, and then restore the original setting before saving the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | TextParagraph SpaceBefore | read paragraph spacing Aspose.Cells | preserve paragraph spacing | LineSpaceSizeType.Points | shape text box spacing | modify SpaceBefore Aspose.Cells | restore original formatting | Aspose.Cells API example | paragraph formatting .NET
// Common Searches: Aspose.Cells get SpaceBefore of TextBox paragraph | change paragraph spacing in Aspose.Cells shape | restore original SpaceBefore after edit Aspose.Cells | C# read and set SpaceBefore points Aspose.Cells | Aspose.Cells TextParagraph spacing example
// Developer Intent: Read a paragraph's SpaceBefore value, apply a temporary change, and then revert to the saved value using Aspose.Cells for .NET.
// Use Cases: Temporarily adjust paragraph spacing for visual emphasis and then revert to the original layout. | Store spacing settings before batch updating multiple shapes to ensure consistent formatting. | Clone a shape, modify its paragraph spacing for a preview, and restore the source shape's spacing afterward.
// AI Prompts: Generate C# code with Aspose.Cells that reads the SpaceBefore of the second paragraph in a TextBox, adds a user‑specified number of points, and then resets it to the original value. | Explain how LineSpaceSizeType.Points and SpaceBeforeSizeType work together to control paragraph spacing in Aspose.Cells. | Provide a script that iterates through all TextParagraphs in a shape, saves each SpaceBefore value, applies a uniform increase, and finally restores the original values.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExample
{
    // Demonstrates how to read the current SpaceBefore value of a TextParagraph in a TextBox shape, modify it using LineSpaceSizeType.Points, and then restore the original setting before saving the workbook with Aspose.Cells for .NET.
    class ReadAndPreserveSpaceBefore
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a text box shape to the worksheet
            Shape shape = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 400, 200);
            shape.Text = "First paragraph\nSecond paragraph";

            // Access the second paragraph (index 1)
            TextParagraph paragraph = shape.TextBody.TextParagraphs[1];

            // Ensure the paragraph uses point units for spacing
            paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;
            paragraph.SpaceBeforeSizeType = LineSpaceSizeType.Points;

            // Read and store the original SpaceBefore value
            double originalSpaceBefore = paragraph.SpaceBefore; // value in points

            // Modify the SpaceBefore value (example: increase by 5 points)
            paragraph.SpaceBefore = originalSpaceBefore + 5;

            // ... perform other operations as needed ...

            // Restore the original SpaceBefore value to preserve settings
            paragraph.SpaceBefore = originalSpaceBefore;

            // Save the workbook
            workbook.Save("ReadAndPreserveSpaceBefore.xlsx");
        }
    }
}
