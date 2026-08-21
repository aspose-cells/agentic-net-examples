// Title: C# – Apply a Dotted‑Grid Pattern Fill to a Rectangle Shape Behind Placeholder Cells with Aspose.Cells
// Description: This example creates a new workbook, writes "Placeholder" into cells A1:C3, adds a rectangle shape that exactly covers that range, sets the shape's FillType to a dotted‑grid pattern (black dots on light‑yellow), sends the shape to the back so the cell text stays visible, and saves the result as PlaceholderCellPatternDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# pattern fill | FillPattern.DottedGrid | rectangle shape behind cells | placeholder cell background | shape Z‑order Aspose.Cells | .NET spreadsheet pattern fill | custom fill pattern Aspose.Cells | cell range background shape | Aspose.Cells shape fill type | C# Aspose.Cells example
// Common Searches: how to add a dotted grid fill to a shape in Aspose.Cells | place a rectangle shape behind cell text Aspose.Cells .NET | Aspose.Cells pattern fill for a cell range | C# Aspose.Cells shape Z‑order back | create placeholder background in Excel with Aspose
// Developer Intent: Add a rectangle shape behind a specific cell range and apply a dotted‑grid pattern fill while keeping the cell content readable.
// Use Cases: Design printable forms where a subtle dotted background marks data‑entry zones without covering labels. | Generate templates that visually indicate where users should insert values, using a patterned backdrop for guidance. | Add decorative or instructional backgrounds to selected ranges in automated reports while preserving text clarity.
// AI Prompts: Generate C# code using Aspose.Cells to cover cells A1:C3 with a rectangle shape, set its FillType to Pattern, choose FillPattern.DottedGrid, define foreground and background colors, and send the shape to the back. | Show how to compute the pixel width and height of a multi‑row, multi‑column range in Aspose.Cells for sizing a shape. | Provide an Aspose.Cells example that creates a placeholder area with a dotted‑grid background while keeping the cell text on top.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a new workbook, writes "Placeholder" into cells A1:C3, adds a rectangle shape that exactly covers that range, sets the shape's FillType to a dotted‑grid pattern (black dots on light‑yellow), sends the shape to the back so the cell text stays visible, and saves the result as PlaceholderCellPatternDemo.xlsx using Aspose.Cells for .NET.
    public class PlaceholderCellPatternDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Fill the range A1:C3 with placeholder text
                for (int row = 0; row < 3; row++)
                {
                    for (int col = 0; col < 3; col++)
                    {
                        sheet.Cells[row, col].PutValue("Placeholder");
                    }
                }

                // Calculate shape size in pixels (width = 3 columns, height = 3 rows)
                int shapeWidth = sheet.Cells.GetColumnWidthPixel(0) * 3;
                int shapeHeight = sheet.Cells.GetRowHeightPixel(0) * 3;

                // Add a rectangle shape that covers the same range as the placeholder cells
                // Overload: AddRectangle(row, column, rowOffset, columnOffset, width, height)
                Shape backgroundShape = sheet.Shapes.AddRectangle(0, 0, 0, 0, shapeWidth, shapeHeight);

                // Position the shape exactly over cells A1:C3
                backgroundShape.Placement = PlacementType.FreeFloating;
                backgroundShape.Top = 0;   // top aligns with row 0
                backgroundShape.Left = 0;  // left aligns with column 0

                // Set the fill type to pattern and define the pattern
                backgroundShape.Fill.FillType = FillType.Pattern;
                backgroundShape.Fill.PatternFill.Pattern = FillPattern.DottedGrid;
                backgroundShape.Fill.PatternFill.ForegroundColor = Color.Black;          // dot color
                backgroundShape.Fill.PatternFill.BackgroundColor = Color.LightYellow;   // background color

                // Send the shape to the back so cell text remains visible
                backgroundShape.ZOrderPosition = 0;

                // Save the workbook
                workbook.Save("PlaceholderCellPatternDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Run error: {ex.Message}");
                throw;
            }
        }
    }
}
