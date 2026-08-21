// Title: Read and Preserve the Original SpaceBefore (Points) of a Shape Paragraph – Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, inserts a text box shape, accesses the second TextParagraph, forces spacing units to points, reads the current SpaceBefore value, stores it, changes the spacing to 12 pt, and saves the file. It shows how to capture and retain the original paragraph spacing before applying custom formatting in Aspose.Cells for .NET.
// Keywords: Aspose.Cells read SpaceBefore | preserve paragraph spacing | Shape TextParagraph points | C# Aspose.Cells spacing | modify SpaceBefore property
// Common Searches: how to get SpaceBefore value from a shape paragraph in Aspose.Cells | preserve original paragraph spacing before changing it C# | read SpaceBefore points from a text box in Aspose.Cells .NET
// Developer Intent: Retrieve a shape paragraph's SpaceBefore (points), keep the original value, then adjust the spacing.
// Use Cases: Log the existing paragraph spacing before applying new formatting for audit trails. | Temporarily modify spacing for a specific report layout while being able to revert to the original settings. | Conditionally adjust paragraph spacing based on the retrieved original SpaceBefore value.
// AI Prompts: Provide C# code that reads the SpaceBefore property of a TextParagraph in Aspose.Cells and saves it before modification. | Show an example that keeps the original paragraph spacing when updating SpaceBefore in a shape's text box using Aspose.Cells for .NET. | Explain how to set spacing units to points before accessing SpaceBefore in Aspose.Cells C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, inserts a text box shape, accesses the second TextParagraph, forces spacing units to points, reads the current SpaceBefore value, stores it, changes the spacing to 12 pt, and saves the file. It shows how to capture and retain the original paragraph spacing before applying custom formatting in Aspose.Cells for .NET.
    public class PreserveSpaceBeforeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                // Parameters: upper left row, upper left column, top offset, left offset, width, height (in pixels)
                Shape shape = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 400, 200);
                shape.Text = "First paragraph\nSecond paragraph";

                // Access the second paragraph (index 1)
                TextParagraph paragraph = shape.TextBody.TextParagraphs[1];

                // Ensure the paragraph uses point units for spacing
                paragraph.SpaceBeforeSizeType = LineSpaceSizeType.Points;
                paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;

                // Preserve original SpaceBefore value
                double originalSpaceBefore = paragraph.SpaceBefore; // value in points
                Console.WriteLine("Original SpaceBefore (points): " + originalSpaceBefore);

                // Modify the SpaceBefore value
                paragraph.SpaceBefore = 12.0; // set new spacing before the paragraph (12 points)

                // Verify the modification
                Console.WriteLine("Modified SpaceBefore (points): " + paragraph.SpaceBefore);

                // Save the workbook to demonstrate persistence
                string outputPath = "PreserveSpaceBeforeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PreserveSpaceBeforeDemo.Run();
        }
    }
}
