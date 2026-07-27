// Title: Concatenate Text from Multiple Child Shapes in Aspose.Cells (C#)
// Description: Shows how to create a workbook, add rectangle shapes as child shapes, read each shape's Text property, concatenate non‑empty values with StringBuilder, display the result, and save the workbook as an XLSX file.
// Keywords: Aspose.Cells | C# | shape text concatenation | child shapes | StringBuilder | Excel workbook | extract shape text | combine rectangle shapes | save XLSX | Aspose.Cells API
// Common Searches: Aspose.Cells concatenate shape text C# | merge text from multiple shapes Aspose.Cells | read Text property of shapes in Excel using Aspose | combine rectangle shape texts in C# | iterate over shapes in Aspose.Cells workbook | save workbook after processing shape texts
// Developer Intent: Combine the Text values of several Shape objects into a single string.
// Use Cases: Build a dynamic header by joining the Text of multiple rectangle shapes before exporting the sheet. | Create a composite label from separate shape parts, concatenate them, and write the result to a cell. | Generate a searchable keyword string by merging annotation texts from all shapes in a worksheet. | Prepare a summary report by aggregating shape captions into one continuous paragraph.
// AI Prompts: Write C# code using Aspose.Cells to concatenate the Text of all shapes in a worksheet and save the workbook. | Explain how to filter shapes by type (e.g., rectangles) before merging their Text values with Aspose.Cells. | Show how to handle null or empty Text properties safely while building a combined string from shape texts.

using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add rectangle shapes as child shapes, read each shape's Text property, concatenate non‑empty values with StringBuilder, display the result, and save the workbook as an XLSX file.
    public class ConcatenateChildShapeTexts
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add individual rectangle shapes (acting as child shapes)
            Shape child1 = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 0, 0, 0, 0, 150, 30);
            child1.Text = "First part ";

            Shape child2 = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 0, 1, 0, 0, 150, 30);
            child2.Text = "Second part ";

            Shape child3 = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 0, 2, 0, 0, 150, 30);
            child3.Text = "Third part";

            // Store the child shapes for easy iteration
            List<Shape> childShapes = new List<Shape> { child1, child2, child3 };

            // Concatenate texts from all child shapes
            StringBuilder concatenatedBuilder = new StringBuilder();
            foreach (Shape child in childShapes)
            {
                if (!string.IsNullOrEmpty(child.Text))
                {
                    concatenatedBuilder.Append(child.Text);
                }
            }

            string concatenatedText = concatenatedBuilder.ToString();
            Console.WriteLine("Concatenated Text: " + concatenatedText);

            // Save the workbook safely
            try
            {
                workbook.Save("ConcatenatedChildShapeTexts.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save workbook: " + saveEx.Message);
            }
        }
    }
}
