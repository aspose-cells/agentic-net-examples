// Title: Proportionally resize an Aspose.Cells shape to fit a target cell range while preserving aspect ratio in C#
// AI Prompts: Determine the pixel dimensions of a target cell area and scale the first worksheet shape to fit within it while maintaining its original aspect ratio using Aspose.Cells for .NET. | Compute the appropriate scaling factor from the target range size and apply proportional width and height changes to a shape in an Excel workbook with Aspose.Cells C#. | After resizing a shape proportionally, align it to the top‑left corner of the specified cell range using Aspose.Cells positioning methods.
// Common Searches: Aspose.Cells C# resize shape to fit inside a specific cell block | maintain shape aspect ratio while scaling in Excel using Aspose.Cells | C# example to place a picture into B2:D5 range with Aspose.Cells | determine cell range dimensions for shape alignment Aspose.Cells .NET | move and size Excel shape based on cell coordinates with Aspose.Cells
// Tags: shape proportional scaling Aspose.Cells | fit shape into cell range C# | preserve aspect ratio Aspose.Cells shape | cell range pixel dimension calculation Aspose.Cells | shape placement using cell offsets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, retrieves the first shape on the first worksheet, calculates the pixel width and height of a target cell range (e.g., B2:D5), scales the shape proportionally to fit inside that range while preserving its aspect ratio, positions the shape at the top‑left corner of the range, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the sheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Get the shape to resize (first shape)
            Shape shape = sheet.Shapes[0];

            // Define the target cell range that the shape should fit into (e.g., B2:D5)
            CellArea targetRange = CellArea.CreateCellArea("B2", "D5");

            // ---------- Calculate pixel dimensions of the target range ----------
            double targetWidth = 0;
            for (int col = targetRange.StartColumn; col <= targetRange.EndColumn; col++)
            {
                targetWidth += sheet.Cells.GetColumnWidthPixel(col);
            }

            double targetHeight = 0;
            for (int row = targetRange.StartRow; row <= targetRange.EndRow; row++)
            {
                targetHeight += sheet.Cells.GetRowHeightPixel(row);
            }

            // ---------- Get current shape size (in pixels) ----------
            int shapeWidthPx = shape.Width;
            int shapeHeightPx = shape.Height;

            // ---------- Compute scaling factor while preserving aspect ratio ----------
            double widthScale = targetWidth / shapeWidthPx;
            double heightScale = targetHeight / shapeHeightPx;
            double scale = Math.Min(widthScale, heightScale); // Fit inside the target range

            // ---------- Apply new size ----------
            shape.Width = (int)(shapeWidthPx * scale);
            shape.Height = (int)(shapeHeightPx * scale);

            // ---------- Position shape at the top‑left corner of the target range ----------
            double leftOffset = 0;
            for (int col = 0; col < targetRange.StartColumn; col++)
            {
                leftOffset += sheet.Cells.GetColumnWidthPixel(col);
            }
            shape.Left = (int)leftOffset;

            double topOffset = 0;
            for (int row = 0; row < targetRange.StartRow; row++)
            {
                topOffset += sheet.Cells.GetRowHeightPixel(row);
            }
            shape.Top = (int)topOffset;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
