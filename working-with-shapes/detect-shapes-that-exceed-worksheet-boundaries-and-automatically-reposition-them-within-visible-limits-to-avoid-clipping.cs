// Title: C# Aspose.Cells: Detect and Reposition Shapes That Exceed Worksheet Used Range to Prevent Clipping
// AI Prompts: Write C# code using Aspose.Cells that iterates over all Shape objects in a worksheet, checks if their UpperLeftRow or UpperLeftColumn are beyond the worksheet's MaxDataRow/MaxDataColumn, and moves the shape back inside the used range. | Create a method that recalculates a shape's Height and Width in points so its bottom‑right corner does not extend past the last used row or column, then apply this adjustment to each shape in the workbook. | Generate a complete console program that loads an Excel file, automatically corrects out‑of‑bounds shapes, and saves the modified workbook.
// Common Searches: aspnet aspose.cells how to move shapes that are outside the used range of an Excel sheet | c# adjust Excel shape size to fit within last used row and column using Aspose.Cells | prevent shape clipping in generated workbook Aspose.Cells .NET | detect shapes positioned beyond worksheet boundaries and reposition them programmatically | auto resize Aspose.Cells Shape objects to stay inside visible area of worksheet
// Tags: detect out-of-bounds shapes Aspose.Cells | reposition Excel shapes within used range .NET | adjust shape dimensions based on worksheet limits | prevent shape clipping in Aspose.Cells workbook | shape coordinate correction Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, determines the maximum used row and column, iterates through every Shape on the first worksheet, moves any shape whose top‑left cell lies outside the used range back inside, and shrinks its height or width when the shape would extend beyond the worksheet limits, then saves the corrected file.
class ShapeRepositioner
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Work with the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine visible limits based on the used range
            int maxRow = Math.Max(0, worksheet.Cells.MaxDataRow);
            int maxColumn = Math.Max(0, worksheet.Cells.MaxDataColumn);

            // Default row height and column width in points (approximate)
            const double defaultRowHeightPoints = 15.0;
            const double defaultColumnWidthPoints = 8.43;

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // ----- Reposition Top‑Left Corner -----
                if (shape.UpperLeftRow > maxRow)
                    shape.UpperLeftRow = maxRow;

                if (shape.UpperLeftColumn > maxColumn)
                    shape.UpperLeftColumn = maxColumn;

                // ----- Adjust Bottom‑Right Corner -----
                int heightInRows = (int)Math.Ceiling(shape.Height / defaultRowHeightPoints);
                int widthInColumns = (int)Math.Ceiling(shape.Width / defaultColumnWidthPoints);

                int bottomRow = shape.UpperLeftRow + heightInRows;
                int rightColumn = shape.UpperLeftColumn + widthInColumns;

                // Shrink height if shape exceeds the used range
                if (bottomRow > maxRow + 1) // +1 because rows are zero‑based
                {
                    int allowedRows = (maxRow + 1) - shape.UpperLeftRow;
                    shape.Height = (int)(allowedRows * defaultRowHeightPoints);
                }

                // Shrink width if shape exceeds the used range
                if (rightColumn > maxColumn + 1) // +1 because columns are zero‑based
                {
                    int allowedColumns = (maxColumn + 1) - shape.UpperLeftColumn;
                    shape.Width = (int)(allowedColumns * defaultColumnWidthPoints);
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
