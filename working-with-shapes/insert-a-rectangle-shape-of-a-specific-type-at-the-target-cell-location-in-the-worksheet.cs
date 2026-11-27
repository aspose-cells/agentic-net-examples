// Title: How to insert a rectangle shape anchored to a specific cell using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add rectangle shape to cell C5 | how to anchor a rectangle shape to a specific worksheet cell using Aspose.Cells | set rectangle height and width in pixels with Aspose.Cells .NET | apply solid fill and border to a shape in Aspose.Cells | save Excel file with inserted rectangle shape using Aspose.Cells
// Tags: Aspose.Cells rectangle shape insertion | cell‑anchored shape Aspose.Cells | shape dimensions pixels Aspose.Cells | shape fill and line styling Aspose.Cells | export workbook with shapes Xlsx

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsRectangleInsert
{
    // The example creates a new workbook, locates cell C5, and adds a rectangle shape anchored to that cell with a height of 80 px and width of 150 px. It sets a solid fill, applies a thin black border, and saves the file as RectangleAtCell.xlsx, handling any runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the resulting workbook will be saved
            const string outputPath = "RectangleAtCell.xlsx";

            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // 3. Define the target cell where the rectangle should be anchored (e.g., "C5")
                const string targetCellName = "C5";
                Cell targetCell = worksheet.Cells[targetCellName];

                // 4. Retrieve zero‑based row and column indices of the target cell
                int topRow = targetCell.Row;
                int leftColumn = targetCell.Column;

                // 5. Define rectangle dimensions (in pixels)
                const int rectangleHeight = 80;
                const int rectangleWidth = 150;
                const int topOffset = 0;   // vertical offset within the row
                const int leftOffset = 0;  // horizontal offset within the column

                // 6. Add the rectangle shape
                RectangleShape rectangle = worksheet.Shapes.AddRectangle(
                    topRow,        // upper‑left row index
                    topOffset,    // vertical offset (pixels)
                    leftColumn,   // upper‑left column index
                    leftOffset,   // horizontal offset (pixels)
                    rectangleHeight,
                    rectangleWidth);

                // 7. Optional appearance customizations
                // Set a solid fill type (default color will be used)
                rectangle.Fill.FillType = FillType.Solid;

                // Add a thin black border
                rectangle.Line.Weight = 1.0f;                     // thickness in points
                rectangle.Line.DashStyle = MsoLineDashStyle.Solid;

                // 8. Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Rectangle shape successfully inserted at cell {targetCellName}.");
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine("An error occurred while inserting the rectangle shape:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
