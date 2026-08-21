// Title: C# – Move a Worksheet Shape to the Back Layer and Align It with a Header Cell using Aspose.Cells
// Description: This Aspose.Cells for .NET example loads an Excel file, ensures a shape exists (adds a rectangle if needed), sends the shape to the back layer with ToFrontOrBack(-1), aligns its top‑left corner to cell A1 using MoveToRange, and saves the modified workbook.
// Keywords: Aspose.Cells | C# | .NET | Excel shape layering | send shape to back | ToFrontOrBack | MoveToRange | align shape with cell | worksheet shape positioning | sample code | GitHub example | watermark shape | logo behind data
// Common Searches: Aspose.Cells move shape to back layer C# | Align shape with cell A1 using Aspose.Cells | How to change Z‑order of Excel shapes in .NET | MoveToRange example Aspose.Cells | Send shape behind worksheet content Aspose.Cells
// Developer Intent: Place a worksheet shape behind all other objects and position it at the header cell.
// Use Cases: Add a company logo as a background element that stays aligned with the top‑left header for every generated report. | Create a watermark shape that sits behind data rows while matching the header cell location. | Automatically re‑order existing shapes in a template so they appear behind content and are anchored to specific header cells.
// AI Prompts: Generate C# code with Aspose.Cells that moves a specific shape to the back layer and aligns it with cell A1. | Show how to iterate over all shapes in a worksheet, send each to the back, and align them with their respective header cells. | Explain the ToFrontOrBack method and demonstrate using MoveToRange to position a shape relative to a cell in an Excel workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeManipulation
{
    // This Aspose.Cells for .NET example loads an Excel file, ensures a shape exists (adds a rectangle if needed), sends the shape to the back layer with ToFrontOrBack(-1), aligns its top‑left corner to cell A1 using MoveToRange, and saves the modified workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "Input.xlsx";
                string outputPath = "Output.xlsx";

                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one shape; create one if none exist
                Shape shape;
                if (worksheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found. Adding a default rectangle shape.");
                    // AddShape(type, upperLeftRow, upperLeftColumn, topOffset, leftOffset, height, width)
                    shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 0, 0, 0, 0, 100, 100);
                }
                else
                {
                    shape = worksheet.Shapes[0];
                }

                // Send the shape to the back layer
                shape.ToFrontOrBack(-1);

                // Header cell position (row 0, column 0)
                int headerRow = 0;
                int headerColumn = 0;

                // Align shape's top‑left corner with the header cell
                shape.MoveToRange(headerRow, headerColumn, headerRow, headerColumn);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
