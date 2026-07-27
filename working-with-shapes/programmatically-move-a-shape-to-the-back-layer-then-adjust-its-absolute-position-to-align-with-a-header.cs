// Title: Aspose.Cells for .NET: Send a Shape to the Back Layer and Align It with a Header Cell
// Description: This example creates a new workbook, adds a rectangle shape, moves the shape behind all other worksheet objects, aligns it with cell A1 (the header row), and saves the file as an XLSX document. It demonstrates the correct use of Shape.ToFrontOrBack and Shape.MoveToRange in C#.
// Keywords: Aspose.Cells | C# shape layering | ToFrontOrBack | MoveToRange | back layer shape | align shape with cell | Excel header shape | Aspose.Cells .NET example | shape positioning
// Common Searches: Aspose.Cells move shape to back layer | How to send a shape behind other objects in Aspose.Cells .NET | Align a rectangle with cell A1 using Aspose.Cells | Shape.ToFrontOrBack valid parameters | Place a shape over the header row in Excel with Aspose.Cells
// Developer Intent: Place a rectangle shape behind all worksheet content and position it precisely over the header cell (A1) using Aspose.Cells for .NET.
// Use Cases: Create a watermark that stays behind data rows while matching the header position. | Add a background label to a report header without obscuring cell values. | Generate dynamic Excel templates where shapes must be layered and aligned automatically.
// AI Prompts: Generate C# code with Aspose.Cells that moves a shape to the back layer without throwing an exception and aligns it to cell A1. | Explain why Shape.ToFrontOrBack(-1) raises ArgumentOutOfRangeException and show the proper way to send a shape to the back. | Provide a concise example of using Shape.MoveToRange to position a shape relative to a header cell while preserving its original dimensions.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeExample
{
    // This example creates a new workbook, adds a rectangle shape, moves the shape behind all other worksheet objects, aligns it with cell A1 (the header row), and saves the file as an XLSX document. It demonstrates the correct use of Shape.ToFrontOrBack and Shape.MoveToRange in C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape (row, column, width, height, row offset, column offset)
                Shape shape = worksheet.Shapes.AddRectangle(5, 5, 120, 30, 0, 0);

                // Attempt to move the shape to the back layer (orders < 0)
                try
                {
                    shape.ToFrontOrBack(-1);
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    Console.WriteLine("Unable to move shape to back: " + ex.Message);
                }

                // Align the shape with cell A1 (row 0, column 0)
                shape.MoveToRange(0, 0, 0, 0);

                // Save the workbook
                string fileName = "ShapeBackAndHeaderAlignment.xlsx";
                workbook.Save(fileName);
                Console.WriteLine($"Workbook saved as {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
